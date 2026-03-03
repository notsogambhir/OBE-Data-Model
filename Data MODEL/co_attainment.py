"""
CO Attainment Calculator
========================
Calculates Course Outcome (CO) attainments from student marks Excel files.
Handles any number of students, COs (2-8+), exams (ST1, ST2, ST3, ETE, ASN, etc.),
and dynamic weightage categories read from the OBE Details sheet.

Usage:
    python co_attainment.py "Input 1.xlsx"
    python co_attainment.py "Input 1.xlsx" "Input 2.xlsx" "Input 3.xlsx"

Each input file is processed independently. Results are saved to
Output_<input_name>.xlsx in the same directory.
"""

import sys
import os
import re
import zipfile
import xml.etree.ElementTree as ET
from collections import OrderedDict

# ─────────────────────────────────────────────────────────────
# SECTION 1: Lightweight XLSX Reader (no pandas/openpyxl needed)
# ─────────────────────────────────────────────────────────────

class XLSXReader:
    """Reads .xlsx files using only the standard library (zipfile + xml)."""

    def __init__(self, path):
        self.path = path
        self.shared_strings = []
        self.sheet_map = OrderedDict()  # {sheet_name: sheet_xml_path}
        self._load_metadata()

    def _load_metadata(self):
        with zipfile.ZipFile(self.path) as z:
            # Shared strings
            if 'xl/sharedStrings.xml' in z.namelist():
                tree = ET.fromstring(z.read('xl/sharedStrings.xml'))
                ns = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                for si in tree.findall('m:si', ns):
                    texts = [t.text for t in si.findall('.//m:t', ns) if t.text is not None]
                    self.shared_strings.append("".join(texts))

            # Workbook sheet list
            wb = ET.fromstring(z.read('xl/workbook.xml'))
            ns = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            for i, sheet in enumerate(wb.find('m:sheets', ns).findall('m:sheet', ns)):
                xml_path = f'xl/worksheets/sheet{i+1}.xml'
                if xml_path in z.namelist():
                    self.sheet_map[sheet.get('name')] = xml_path

    @property
    def sheet_names(self):
        return list(self.sheet_map.keys())

    def _col_index(self, ref):
        """Convert cell reference like 'AB12' to 0-based column index."""
        col = ''
        for ch in ref:
            if ch.isalpha():
                col += ch
            else:
                break
        idx = 0
        for c in col.upper():
            idx = idx * 26 + (ord(c) - ord('A') + 1)
        return idx - 1

    def read_sheet(self, sheet_name):
        """Return list-of-lists for the given sheet, with column gaps filled."""
        ns = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        with zipfile.ZipFile(self.path) as z:
            tree = ET.fromstring(z.read(self.sheet_map[sheet_name]))
        sd = tree.find('m:sheetData', ns)
        rows = []
        for row_el in sd.findall('m:row', ns):
            row = []
            for c in row_el.findall('m:c', ns):
                ref = c.get('r', '')
                col_idx = self._col_index(ref)
                # Pad with empty strings if there are gaps
                while len(row) < col_idx:
                    row.append('')
                v = c.find('m:v', ns)
                if v is not None and v.text is not None:
                    if c.get('t') == 's':
                        row.append(self.shared_strings[int(v.text)])
                    else:
                        row.append(v.text)
                else:
                    row.append('')
            rows.append(row)
        return rows


# ─────────────────────────────────────────────────────────────
# SECTION 2: Parse OBE Details (Thresholds, Weights, Levels)
# ─────────────────────────────────────────────────────────────

def parse_obe_details(reader):
    """
    Reads the first sheet ('OBE Details') and extracts:
      - target_pct: the course target percentage (e.g. 60)
      - weights: dict mapping category keyword -> weight
                 e.g. {'Internal': 0.4, 'External': 0.6}
                 or   {'Internal': 0.3, 'External': 0.6, 'Assignment': 0.1}
      - levels: list of (level_number, threshold) sorted desc by level
                 e.g. [(3, 0.8), (2, 0.7), (1, 0.6)]
    """
    rows = reader.read_sheet(reader.sheet_names[0])

    target_pct = 60.0
    weights = OrderedDict()
    levels = []

    mode = None  # 'weights' or 'levels'
    for row in rows:
        if len(row) < 2:
            continue
        c0 = str(row[0]).strip()
        c1 = str(row[1]).strip()

        # Look for Threshold row
        if c0.lower() == 'threshold' and c1:
            try:
                target_pct = float(c1)
            except:
                pass
            continue

        # Detect section headers
        if c0.lower() == 'types' and c1.lower() == 'weightages':
            mode = 'weights'
            continue
        if 'co score' in c0.lower() and 'student' in c1.lower():
            mode = 'levels'
            continue

        if mode == 'weights' and c0 and c1:
            try:
                w = float(c1)
            except:
                continue
            # Determine the category keyword
            c0_lower = c0.lower()
            if 'internal' in c0_lower or 'avg' in c0_lower or 'st1' in c0_lower:
                weights['Internal'] = w
            elif 'external' in c0_lower or 'ete' in c0_lower:
                weights['External'] = w
            elif 'assign' in c0_lower or 'asn' in c0_lower:
                weights['Assignment'] = w
            else:
                weights[c0] = w

        if mode == 'levels' and c0 and c1:
            try:
                lv = int(float(c0))
                th = float(c1)
                levels.append((lv, th))
            except:
                pass

    # Sort levels descending so we match highest first
    levels.sort(key=lambda x: x[0], reverse=True)

    # Fallback defaults
    if not weights:
        weights = OrderedDict([('Internal', 0.4), ('External', 0.6)])
    if not levels:
        levels = [(3, 0.8), (2, 0.7), (1, 0.6)]

    return target_pct, weights, levels


# ─────────────────────────────────────────────────────────────
# SECTION 3: Discover Exams (Mapping + Result sheet pairs)
# ─────────────────────────────────────────────────────────────

def discover_exams(reader):
    """
    Scans sheet names for matching pairs:
      '<EXAM> Ques Mapping' / '<EXAM> Mapping' + '<EXAM> Result'
    Returns a list of dicts: {name, map_sheet, res_sheet, category}
    where category is 'Internal', 'External', or 'Assignment'.
    """
    sheets = reader.sheet_names
    exams = []
    seen = set()

    # Find all mapping sheets
    for s in sheets:
        m = None
        if 'Mapping' in s:
            # Extract exam name (everything before ' Ques Mapping' or ' Mapping')
            m = s.replace(' Ques Mapping', '').replace(' Mapping', '').strip()
        if m and m not in seen:
            # Find corresponding result sheet
            res = None
            for s2 in sheets:
                if s2.startswith(m) and 'Result' in s2:
                    res = s2
                    break
            if res:
                # Categorise
                m_upper = m.upper()
                if m_upper.startswith('ETE') or 'EXTERNAL' in m_upper:
                    cat = 'External'
                elif m_upper.startswith('ASN') or 'ASSIGN' in m_upper:
                    cat = 'Assignment'
                else:
                    cat = 'Internal'  # ST1, ST2, ST3, FA, etc.
                exams.append({'name': m, 'map_sheet': s, 'res_sheet': res, 'category': cat})
                seen.add(m)

    return exams


# ─────────────────────────────────────────────────────────────
# SECTION 4: Parse Mapping & Result sheets
# ─────────────────────────────────────────────────────────────

def parse_mapping(reader, sheet_name):
    """
    Returns a list of dicts, each representing a question:
      {'q_id': str, 'max_marks': float, 'cos': {<CO_name>: 1/0, ...}}
    Only includes questions with max_marks > 0 and at least one CO mapped.
    """
    rows = reader.read_sheet(sheet_name)
    if not rows:
        return [], []

    header = [str(h).strip() for h in rows[0]]

    # Find Q_Id column (first col) and Max Marks (second col)
    # Find CO columns (any column whose header matches CO\d+)
    co_cols = []
    for i, h in enumerate(header):
        if re.match(r'^CO\d+$', h, re.IGNORECASE):
            co_cols.append((i, h.upper()))

    questions = []
    for row in rows[1:]:
        if len(row) < 2:
            continue
        q_id = str(row[0]).strip()
        if not q_id:
            continue
        try:
            max_m = float(row[1])
        except:
            continue
        if max_m <= 0:
            continue

        cos = {}
        mapped_any = False
        for ci, cname in co_cols:
            if ci < len(row):
                try:
                    val = int(float(row[ci]))
                except:
                    val = 0
                cos[cname] = val
                if val > 0:
                    mapped_any = True
            else:
                cos[cname] = 0

        if mapped_any:
            questions.append({'q_id': q_id, 'max_marks': max_m, 'cos': cos})

    all_cos = sorted(set(c for _, c in co_cols))
    return questions, all_cos


def parse_results(reader, sheet_name):
    """
    Returns a list of student dicts:
      {'roll': str, 'name': str, 'marks': {<q_id>: float or None}}
    None means unattempted.
    """
    rows = reader.read_sheet(sheet_name)
    if not rows:
        return []

    header = [str(h).strip() for h in rows[0]]

    # Find roll number and name columns
    roll_idx = None
    name_idx = None
    for i, h in enumerate(header):
        hl = h.lower()
        if 'admission' in hl or 'roll' in hl:
            roll_idx = i
        if 'name' in hl and 'student' in hl:
            name_idx = i

    if roll_idx is None:
        return []

    # Find question columns: everything after the metadata columns and before
    # 'Total Marks', 'max Marks', 'Maximum marks' etc.
    # We identify the first question column and the last
    meta_keywords = {'sr.no', 'sr no', 'admission', 'roll', 'name', 'course', 'exam', 'total', 'max', 'maximum'}

    q_cols = []  # (index, header_name)
    for i, h in enumerate(header):
        hl = h.lower().strip()
        is_meta = False
        for kw in meta_keywords:
            if kw in hl:
                is_meta = True
                break
        if not is_meta and h.strip():
            q_cols.append((i, h.strip()))

    students = []
    for row in rows[1:]:
        if len(row) <= max(roll_idx, 1):
            continue
        roll = str(row[roll_idx]).strip().replace('\xa0', '').strip()
        if not roll:
            continue
        name = str(row[name_idx]).strip().replace('\xa0', '').strip() if name_idx is not None and name_idx < len(row) else ''

        marks = {}
        for ci, qname in q_cols:
            if ci < len(row):
                val = str(row[ci]).strip().upper()
                if val == 'U' or val == '' or val == 'AB':
                    marks[qname] = None  # Unattempted
                else:
                    try:
                        marks[qname] = float(val)
                    except:
                        marks[qname] = None
            else:
                marks[qname] = None

        students.append({'roll': roll, 'name': name, 'marks': marks})

    return students


# ─────────────────────────────────────────────────────────────
# SECTION 5: Core CO Calculation per Student per Exam
# ─────────────────────────────────────────────────────────────

def calc_student_co_pct(student_marks, questions, co_name):
    """
    For a single student in a single exam, compute the CO percentage.
    - Only considers questions mapped to co_name.
    - Skips unattempted questions (None) from BOTH numerator and denominator.
    Returns float percentage or None if no valid questions.
    """
    obtained = 0.0
    max_total = 0.0

    for q in questions:
        if q['cos'].get(co_name, 0) <= 0:
            continue
        val = student_marks.get(q['q_id'])
        if val is None:
            continue  # Unattempted – skip both num and denom
        obtained += val
        max_total += q['max_marks']

    if max_total == 0:
        return None
    return (obtained / max_total) * 100.0


# ─────────────────────────────────────────────────────────────
# SECTION 6: Main Processing Pipeline
# ─────────────────────────────────────────────────────────────

def process_file(filepath):
    """
    End-to-end processing of one input Excel file.
    Returns (student_details, attainment_summary, meta_info).
    """
    print(f"\n{'='*60}")
    print(f"  Processing: {os.path.basename(filepath)}")
    print(f"{'='*60}")

    reader = XLSXReader(filepath)

    # --- Step 1: Parse OBE Details ---
    target_pct, weights, levels = parse_obe_details(reader)
    print(f"  Target: {target_pct}%")
    print(f"  Weights: {dict(weights)}")
    print(f"  Levels: {levels}")

    # --- Step 2: Discover exams ---
    exams = discover_exams(reader)
    print(f"  Exams found: {[e['name'] + ' (' + e['category'] + ')' for e in exams]}")

    # --- Step 3: Parse mapping + results for each exam ---
    exam_data = []
    all_cos_global = set()
    all_students = OrderedDict()  # roll -> name

    for exam in exams:
        questions, cos = parse_mapping(reader, exam['map_sheet'])
        students = parse_results(reader, exam['res_sheet'])

        # Skip exams with no valid mapped questions
        if not questions:
            print(f"    Skipping {exam['name']}: no valid question mappings (all zero marks or unmapped)")
            continue

        for s in students:
            if s['roll'] not in all_students:
                all_students[s['roll']] = s['name']

        all_cos_global.update(cos)
        exam_data.append({
            'name': exam['name'],
            'category': exam['category'],
            'questions': questions,
            'cos': cos,
            'students': {s['roll']: s for s in students}
        })

    all_cos_sorted = sorted(all_cos_global)
    print(f"  All COs across exams: {all_cos_sorted}")
    print(f"  Total students: {len(all_students)}")

    # --- Step 4: Compute per-student, per-exam, per-CO percentages ---
    # Structure: student_exam_co[roll][exam_name][CO] = pct or None
    student_exam_co = {}
    for roll in all_students:
        student_exam_co[roll] = {}
        for ed in exam_data:
            s_data = ed['students'].get(roll)
            if s_data is None:
                student_exam_co[roll][ed['name']] = {co: None for co in all_cos_sorted}
            else:
                co_pcts = {}
                for co in all_cos_sorted:
                    co_pcts[co] = calc_student_co_pct(s_data['marks'], ed['questions'], co)
                student_exam_co[roll][ed['name']] = co_pcts

    # --- Step 5: Compute weighted final CO% per student ---
    # Group exams by category
    cat_exams = {}
    for ed in exam_data:
        cat_exams.setdefault(ed['category'], []).append(ed['name'])

    student_final_co = {}  # roll -> {CO: weighted_pct}
    student_details = {}   # roll -> {CO: {exam_pcts..., cat_avgs..., final}}

    for roll in all_students:
        student_final_co[roll] = {}
        student_details[roll] = {'name': all_students[roll]}

        for co in all_cos_sorted:
            detail = {}
            weighted_total = 0.0
            weight_sum = 0.0

            for cat, w in weights.items():
                cat_exam_names = cat_exams.get(cat, [])
                if not cat_exam_names:
                    continue

                # Collect this student's CO% across all exams in this category
                pcts = []
                for ename in cat_exam_names:
                    p = student_exam_co[roll].get(ename, {}).get(co)
                    detail[f"{ename}_{co}%"] = round(p, 2) if p is not None else 'N/A'
                    if p is not None:
                        pcts.append(p)

                if pcts:
                    cat_avg = sum(pcts) / len(pcts)
                    detail[f"{cat}_Avg"] = round(cat_avg, 2)
                    weighted_total += cat_avg * w
                    weight_sum += w
                else:
                    detail[f"{cat}_Avg"] = 'N/A'

            if weight_sum > 0:
                # Normalize by sum of weights that actually had data
                # This handles cases where a student misses an entire category
                final_pct = weighted_total / weight_sum * (sum(weights.values()) / 1.0)
                # Actually, simpler: if a student has no data for a category,
                # treat that category's contribution as 0 (penalty approach)
                final_pct = weighted_total
                detail['Final_%'] = round(final_pct, 2)
                student_final_co[roll][co] = final_pct
            else:
                detail['Final_%'] = 'N/A'
                student_final_co[roll][co] = None

            student_details[roll][co] = detail

    # --- Step 6: Class-level aggregation and attainment levels ---
    attainment_summary = OrderedDict()
    for co in all_cos_sorted:
        valid_count = 0
        success_count = 0
        for roll in all_students:
            pct = student_final_co[roll].get(co)
            if pct is not None:
                valid_count += 1
                if pct >= target_pct:
                    success_count += 1

        if valid_count > 0:
            success_rate = success_count / valid_count
        else:
            success_rate = 0.0

        # Determine level
        level = 0
        for lv, th in levels:
            if success_rate >= th:
                level = lv
                break

        attainment_summary[co] = {
            'Students_Attempted': valid_count,
            'Students_Meeting_Target': success_count,
            'Success_Rate_%': round(success_rate * 100, 2),
            'Attainment_Level': level
        }

    # --- Print summary ---
    print(f"\n  --- CO Attainment Summary ---")
    print(f"  {'CO':<8} {'Attempted':>10} {'Met Target':>12} {'Success%':>10} {'Level':>6}")
    print(f"  {'-'*50}")
    for co, info in attainment_summary.items():
        print(f"  {co:<8} {info['Students_Attempted']:>10} {info['Students_Meeting_Target']:>12} {info['Success_Rate_%']:>9.2f}% {info['Attainment_Level']:>6}")

    meta_info = {
        'target_pct': target_pct,
        'weights': dict(weights),
        'levels': levels,
        'exams': [{'name': e['name'], 'category': e['category']} for e in exam_data],
        'all_cos': all_cos_sorted,
        'total_students': len(all_students)
    }

    return all_students, student_final_co, student_details, attainment_summary, meta_info


# ─────────────────────────────────────────────────────────────
# SECTION 7: Excel Output Writer (pure zipfile, no openpyxl)
# ─────────────────────────────────────────────────────────────

def write_output_xlsx(filepath, all_students, student_final_co, attainment_summary, meta_info):
    """Write results to a simple .xlsx file using zipfile + XML."""

    all_cos = meta_info['all_cos']

    # Build shared strings
    ss = []
    ss_map = {}

    def add_ss(s):
        s = str(s)
        if s not in ss_map:
            ss_map[s] = len(ss)
            ss.append(s)
        return ss_map[s]

    # Pre-populate shared strings
    sheet1_headers = ['Roll No.', 'Student Name'] + [f'{co} %' for co in all_cos] + [f'{co} Target Met' for co in all_cos]
    for h in sheet1_headers:
        add_ss(h)

    sheet2_headers = ['Course Outcome', 'Students Attempted', 'Students Meeting Target', 'Success Rate %', 'Attainment Level']
    for h in sheet2_headers:
        add_ss(h)

    for roll, name in all_students.items():
        add_ss(roll)
        add_ss(name)

    for co in all_cos:
        add_ss(co)
    add_ss('Yes')
    add_ss('No')
    add_ss('N/A')

    # Build meta strings
    meta_headers = ['Parameter', 'Value']
    for h in meta_headers:
        add_ss(h)
    add_ss('Target Percentage')
    add_ss(str(meta_info['target_pct']))
    for cat, w in meta_info['weights'].items():
        add_ss(f'Weight: {cat}')
        add_ss(str(w))
    for lv, th in meta_info['levels']:
        add_ss(f'Level {lv} Threshold')
        add_ss(str(th))

    def col_letter(idx):
        """0-based index to Excel column letter(s)."""
        result = ''
        while True:
            result = chr(idx % 26 + ord('A')) + result
            idx = idx // 26 - 1
            if idx < 0:
                break
        return result

    def make_row_xml(row_num, cells):
        """cells is list of (value, is_string)."""
        parts = [f'<row r="{row_num}">']
        for ci, (val, is_str) in enumerate(cells):
            ref = f'{col_letter(ci)}{row_num}'
            if is_str:
                si = add_ss(str(val))
                parts.append(f'<c r="{ref}" t="s"><v>{si}</v></c>')
            else:
                parts.append(f'<c r="{ref}"><v>{val}</v></c>')
        parts.append('</row>')
        return ''.join(parts)

    # ---- Sheet 1: Student Details ----
    s1_rows = []
    # Header
    cells = [(h, True) for h in sheet1_headers]
    s1_rows.append(make_row_xml(1, cells))

    row_num = 2
    target = meta_info['target_pct']
    for roll, name in all_students.items():
        cells = [(roll, True), (name, True)]
        for co in all_cos:
            pct = student_final_co[roll].get(co)
            if pct is not None:
                cells.append((round(pct, 2), False))
            else:
                cells.append(('N/A', True))
        for co in all_cos:
            pct = student_final_co[roll].get(co)
            if pct is not None:
                cells.append(('Yes' if pct >= target else 'No', True))
            else:
                cells.append(('N/A', True))
        s1_rows.append(make_row_xml(row_num, cells))
        row_num += 1

    # ---- Sheet 2: Attainment Summary ----
    s2_rows = []
    cells = [(h, True) for h in sheet2_headers]
    s2_rows.append(make_row_xml(1, cells))

    row_num = 2
    for co, info in attainment_summary.items():
        cells = [
            (co, True),
            (info['Students_Attempted'], False),
            (info['Students_Meeting_Target'], False),
            (info['Success_Rate_%'], False),
            (info['Attainment_Level'], False),
        ]
        s2_rows.append(make_row_xml(row_num, cells))
        row_num += 1

    # ---- Sheet 3: Configuration ----
    s3_rows = []
    cells = [(h, True) for h in meta_headers]
    s3_rows.append(make_row_xml(1, cells))

    row_num = 2
    cfg_items = [('Target Percentage', str(meta_info['target_pct']))]
    for cat, w in meta_info['weights'].items():
        cfg_items.append((f'Weight: {cat}', str(w)))
    for lv, th in meta_info['levels']:
        cfg_items.append((f'Level {lv} Threshold', str(th)))
    for k, v in cfg_items:
        cells = [(k, True), (v, True)]
        s3_rows.append(make_row_xml(row_num, cells))
        row_num += 1

    # ---- Assemble XLSX ----
    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    rel_ns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    ct_ns = 'http://schemas.openxmlformats.org/package/2006/content-types'
    pkg_rel_ns = 'http://schemas.openxmlformats.org/package/2006/relationships'

    shared_strings_xml = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="{ns}" count="{len(ss)}" uniqueCount="{len(ss)}">'
    for s in ss:
        escaped = s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        shared_strings_xml += f'<si><t>{escaped}</t></si>'
    shared_strings_xml += '</sst>'

    def make_sheet_xml(rows_xml):
        return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{ns}"><sheetData>{"".join(rows_xml)}</sheetData></worksheet>'

    workbook_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="{ns}" xmlns:r="{rel_ns}">
<sheets>
<sheet name="Student Details" sheetId="1" r:id="rId1"/>
<sheet name="Attainment Summary" sheetId="2" r:id="rId2"/>
<sheet name="Configuration" sheetId="3" r:id="rId3"/>
</sheets>
</workbook>'''

    wb_rels = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{pkg_rel_ns}">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>'''

    root_rels = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{pkg_rel_ns}">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''

    content_types = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="{ct_ns}">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>'''

    with zipfile.ZipFile(filepath, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', content_types)
        zf.writestr('_rels/.rels', root_rels)
        zf.writestr('xl/workbook.xml', workbook_xml)
        zf.writestr('xl/_rels/workbook.xml.rels', wb_rels)
        zf.writestr('xl/sharedStrings.xml', shared_strings_xml)
        zf.writestr('xl/worksheets/sheet1.xml', make_sheet_xml(s1_rows))
        zf.writestr('xl/worksheets/sheet2.xml', make_sheet_xml(s2_rows))
        zf.writestr('xl/worksheets/sheet3.xml', make_sheet_xml(s3_rows))

    print(f"\n  Output saved to: {filepath}")


# ─────────────────────────────────────────────────────────────
# SECTION 8: Main Entry Point
# ─────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        # Default: process all Input *.xlsx in current directory
        files = sorted([f for f in os.listdir('.') if f.lower().startswith('input') and f.lower().endswith('.xlsx')])
        if not files:
            print("Usage: python co_attainment.py <file1.xlsx> [file2.xlsx] ...")
            sys.exit(1)
    else:
        files = sys.argv[1:]

    for filepath in files:
        if not os.path.exists(filepath):
            print(f"ERROR: File not found: {filepath}")
            continue

        try:
            all_students, student_final_co, student_details, attainment_summary, meta_info = process_file(filepath)

            # Generate output filename
            base = os.path.splitext(os.path.basename(filepath))[0]
            out_path = os.path.join(os.path.dirname(filepath) or '.', f"Output_{base}.xlsx")
            write_output_xlsx(out_path, all_students, student_final_co, attainment_summary, meta_info)
        except Exception as e:
            print(f"ERROR processing {filepath}: {e}")
            import traceback
            traceback.print_exc()

    print(f"\n{'='*60}")
    print("  All files processed!")
    print(f"{'='*60}")


if __name__ == '__main__':
    main()
