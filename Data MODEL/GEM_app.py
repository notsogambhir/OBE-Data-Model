"""
CO Attainment Calculator — Standalone Application
=====================================================
Calculates how well students achieved Course Outcomes (COs).
Includes a graphical user interface (GUI) for ease of use.

Install: pip install pandas openpyxl
Run:     python GEM_app.py
"""

import pandas as pd
import numpy as np
import os
import re
import threading
import tkinter as tk
from tkinter import filedialog, messagebox


# ─────────────────────────────────────────────
# CALCULATION LOGIC
# ─────────────────────────────────────────────

def calculate_co_attainment(input_file, output_file=None):
    """Process one Excel file and calculate CO attainment."""
    print(f"\n{'='*60}")
    print(f"  Processing: {input_file}")
    print(f"{'='*60}")

    excel = pd.ExcelFile(input_file)

    # --- STEP 1: Read settings from OBE Details sheet ---
    obe = excel.parse(0, header=None)
    target = 60.0
    weights = {}
    levels = []
    reading = None

    for i in range(len(obe)):
        a = str(obe.iloc[i, 0]).strip() if pd.notna(obe.iloc[i, 0]) else ''
        b = str(obe.iloc[i, 1]).strip() if obe.shape[1] > 1 and pd.notna(obe.iloc[i, 1]) else ''

        if a.lower() == 'threshold' and b:
            try:
                target = float(b)
            except:
                pass

        if a.lower() == 'types' and 'weightage' in b.lower():
            reading = 'weights'
            continue

        if 'co score' in a.lower() and 'student' in b.lower():
            reading = 'levels'
            continue

        if reading == 'weights' and a and b:
            try:
                w = float(b)
                if 'internal' in a.lower():
                    weights['Internal'] = w
                elif 'external' in a.lower() or 'ete' in a.lower():
                    weights['External'] = w
                elif 'assign' in a.lower():
                    weights['Assignment'] = w
            except:
                pass

        if reading == 'levels' and a and b:
            try:
                levels.append((int(float(a)), float(b)))
            except:
                pass

    if not weights:
        weights = {'Internal': 0.4, 'External': 0.6}
    if not levels:
        levels = [(3, 0.8), (2, 0.7), (1, 0.6)]
    levels.sort(reverse=True)

    print(f"  Target: {target}%")
    print(f"  Weights: {weights}")

    # --- STEP 2: Find all exams in the workbook ---
    exams = []
    found = set()

    for sheet in excel.sheet_names:
        if 'Mapping' not in sheet:
            continue

        name = sheet.replace(' Ques Mapping', '').replace(' Mapping', '').strip()
        if name in found:
            continue

        result_sheet = None
        for s in excel.sheet_names:
            if s.startswith(name) and 'Result' in s:
                result_sheet = s
                break

        if result_sheet:
            if 'ETE' in name.upper():
                cat = 'External'
            elif 'ASN' in name.upper():
                cat = 'Assignment'
            else:
                cat = 'Internal'

            exams.append({'name': name, 'mapping': sheet, 'result': result_sheet, 'category': cat})
            found.add(name)

    print(f"  Exams found: {[e['name'] for e in exams]}")

    # --- STEP 3: Calculate CO% per student for each exam ---
    all_students = {}
    all_cos = set()
    exam_scores = []

    for exam in exams:
        mapping = excel.parse(exam['mapping'])
        mapping.columns = [str(c).strip() for c in mapping.columns]

        co_columns = [c for c in mapping.columns if re.match(r'^CO\d+$', c, re.IGNORECASE)]
        co_columns = sorted(co_columns, key=lambda x: int(re.search(r'\d+', x).group()))

        if not co_columns:
            continue

        for c in co_columns:
            mapping.rename(columns={c: c.upper()}, inplace=True)
        co_columns = [c.upper() for c in co_columns]

        q_col = mapping.columns[0]
        mm_col = mapping.columns[1]

        mapping[q_col] = mapping[q_col].astype(str).str.strip()
        mapping[mm_col] = pd.to_numeric(mapping[mm_col], errors='coerce').fillna(0)

        for co in co_columns:
            mapping[co] = mapping[co].apply(lambda v: int(v) if isinstance(v, bool) else v)
            mapping[co] = pd.to_numeric(mapping[co], errors='coerce').fillna(0).astype(int)

        mapping = mapping[mapping[mm_col] > 0]
        mapping = mapping[mapping[co_columns].sum(axis=1) > 0]

        if mapping.empty:
            continue

        results = excel.parse(exam['result'])
        results.columns = [str(c).strip() for c in results.columns]

        roll_col = next((c for c in results.columns if 'admission' in c.lower() or 'roll' in c.lower()), None)
        name_col = next((c for c in results.columns if 'name' in c.lower()), None)

        if not roll_col:
            continue

        results[roll_col] = results[roll_col].astype(str).str.strip()
        results = results[results[roll_col].str.len() > 0]

        for _, row in results.iterrows():
            roll = row[roll_col]
            name = str(row[name_col]).strip() if name_col else ''
            if roll not in all_students:
                all_students[roll] = name

        scores = pd.DataFrame()
        scores['Roll'] = results[roll_col].values

        for co in co_columns:
            linked_questions = mapping[mapping[co] > 0][q_col].tolist()
            max_marks_map = mapping.set_index(q_col)[mm_col]

            if not linked_questions:
                continue

            co_percentages = []
            for _, student in results.iterrows():
                obtained = 0
                maximum = 0

                for q in linked_questions:
                    if q not in results.columns:
                        continue
                    mark = student[q]
                    if pd.isna(mark) or str(mark).strip().upper() in ('U', '', 'AB'):
                        continue
                    try:
                        obtained += float(mark)
                        maximum += max_marks_map.get(q, 0)
                    except (ValueError, TypeError):
                        continue

                co_percentages.append((obtained / maximum * 100) if maximum > 0 else np.nan)

            scores[co] = co_percentages
            all_cos.add(co)

        exam_scores.append({
            'name': exam['name'],
            'category': exam['category'],
            'scores': scores.set_index('Roll')
        })

    all_cos = sorted(all_cos)

    # --- STEP 4: Combine exams using weighted average ---
    category_exams = {}
    for es in exam_scores:
        category_exams.setdefault(es['category'], []).append(es)

    rolls = list(all_students.keys())
    names = list(all_students.values())
    final = pd.DataFrame({'Roll': rolls, 'Name': names}).set_index('Roll')

    for co in all_cos:
        weighted_total = pd.Series(0.0, index=final.index)
        for category, weight in weights.items():
            exams_in_cat = category_exams.get(category, [])
            if not exams_in_cat:
                continue
            pcts = pd.DataFrame(index=final.index)
            for es in exams_in_cat:
                if co in es['scores'].columns:
                    pcts[es['name']] = es['scores'][co].reindex(final.index)
            if not pcts.empty:
                avg = pcts.mean(axis=1, skipna=True)
                weighted_total += avg.fillna(0) * weight

        final[f'{co} %'] = weighted_total.replace(0, np.nan).round(2)
        final[f'{co} Met Target'] = final[f'{co} %'].apply(
            lambda x: 'Yes' if pd.notna(x) and x >= target else ('No' if pd.notna(x) else 'N/A')
        )

    # --- STEP 5: Calculate class-level attainment ---
    summary = []
    for co in all_cos:
        valid_scores = final[f'{co} %'].dropna()
        total_students = len(valid_scores)
        students_passed = int((valid_scores >= target).sum())
        success_rate = (students_passed / total_students) if total_students > 0 else 0

        level = 0
        for lv, threshold in levels:
            if success_rate >= threshold:
                level = lv
                break
        summary.append({
            'CO': co,
            'Students Attempted': total_students,
            'Met Target': students_passed,
            'Success Rate %': round(success_rate * 100, 2),
            'Attainment Level': level
        })

    summary_df = pd.DataFrame(summary)

    # --- STEP 6: Save results to Excel ---
    if output_file is None:
        output_file = os.path.join(os.path.dirname(input_file) or '.', f"Output_{os.path.basename(input_file)}")

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        final.reset_index().to_excel(writer, sheet_name='Student Details', index=False)
        summary_df.to_excel(writer, sheet_name='Attainment Summary', index=False)

    print(f"  Saved: {output_file}")


# ─────────────────────────────────────────────
# GRAPHICAL USER INTERFACE (GUI)
# ─────────────────────────────────────────────

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("CO Attainment Calculator")
        self.root.geometry("600x300")
        self.root.resizable(False, False)

        # Variables
        self.input_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.status_var = tk.StringVar()
        self.status_var.set("Ready.")

        self.setup_ui()

    def setup_ui(self):
        frame = tk.Frame(self.root, padx=20, pady=20)
        frame.pack(fill=tk.BOTH, expand=True)

        lbl_title = tk.Label(frame, text="Course Outcome (CO) Attainment Calculator", font=("Arial", 14, "bold"))
        lbl_title.grid(row=0, column=0, columnspan=3, pady=(0, 20))

        # Input File
        tk.Label(frame, text="Input File:", font=("Arial", 10)).grid(row=1, column=0, sticky="w", pady=5)
        tk.Entry(frame, textvariable=self.input_var, width=50, font=("Arial", 10)).grid(row=1, column=1, padx=10, pady=5)
        tk.Button(frame, text="Browse...", command=self.browse_input, width=10).grid(row=1, column=2, pady=5)

        # Output File
        tk.Label(frame, text="Save As:", font=("Arial", 10)).grid(row=2, column=0, sticky="w", pady=5)
        tk.Entry(frame, textvariable=self.output_var, width=50, font=("Arial", 10)).grid(row=2, column=1, padx=10, pady=5)
        tk.Button(frame, text="Browse...", command=self.browse_output, width=10).grid(row=2, column=2, pady=5)

        # Run Button
        self.btn_run = tk.Button(frame, text="Run Calculation", command=self.run_calc, font=("Arial", 12, "bold"), 
                                 bg="#4CAF50", fg="white", activebackground="#45a049", activeforeground="white", width=20, height=2)
        self.btn_run.grid(row=3, column=0, columnspan=3, pady=25)

        # Status Bar
        tk.Label(frame, textvariable=self.status_var, font=("Arial", 9, "italic"), fg="gray").grid(row=4, column=0, columnspan=3, sticky="w")

    def browse_input(self):
        path = filedialog.askopenfilename(title="Select Input Excel File", filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.input_var.set(path)
            # Auto-suggest output name
            filename = os.path.basename(path)
            suggested_out = os.path.join(os.path.dirname(path), f"Output_{filename}")
            self.output_var.set(suggested_out)

    def browse_output(self):
        initial_dir, initial_file = "", ""
        if self.input_var.get():
            initial_dir = os.path.dirname(self.input_var.get())
            initial_file = f"Output_{os.path.basename(self.input_var.get())}"
            
        path = filedialog.asksaveasfilename(
            title="Save Output Excel File As",
            initialdir=initial_dir,
            initialfile=initial_file,
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if path:
            self.output_var.set(path)

    def run_calc(self):
        input_file = self.input_var.get().strip()
        output_file = self.output_var.get().strip()
        
        if not input_file or not os.path.exists(input_file):
            messagebox.showerror("Error", "Please select a valid input Excel file.")
            return
        if not output_file:
            messagebox.showerror("Error", "Please select an output save location.")
            return

        self.btn_run.config(state=tk.DISABLED, text="Processing...")
        self.status_var.set("Running calculation... Please wait.")
        
        # Run in thread so GUI doesn't freeze
        threading.Thread(target=self.process_thread, args=(input_file, output_file), daemon=True).start()

    def process_thread(self, input_file, output_file):
        try:
            calculate_co_attainment(input_file, output_file)
            self.root.after(0, self.on_success, output_file)
        except Exception as e:
            self.root.after(0, self.on_error, str(e))

    def on_success(self, output_file):
        self.status_var.set("✅ Complete! Saved successfully.")
        self.btn_run.config(state=tk.NORMAL, text="Run Calculation")
        messagebox.showinfo("Success", f"Calculation complete!\nSaved to:\n{output_file}")

    def on_error(self, err_msg):
        self.status_var.set("❌ Error occurred.")
        self.btn_run.config(state=tk.NORMAL, text="Run Calculation")
        messagebox.showerror("Error", f"Failed to process file:\n{err_msg}")


# ─────────────────────────────────────────────
# MAIN ENTRY POINT
# ─────────────────────────────────────────────

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
