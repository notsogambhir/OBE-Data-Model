"""
GEM_quant.py — Vectorized, minimal 'quant-style' CO Attainment Calculator.
Dense, highly efficient, and compact (< 90 lines of code).
"""
import pandas as pd, numpy as np, tkinter as tk, os
from tkinter import filedialog, messagebox

def calc_co(in_file, out_file=None):
    xl = pd.ExcelFile(in_file)
    out = out_file or os.path.join(os.path.dirname(in_file), f"Output_{os.path.basename(in_file)}")
    
    # 1. Parse Settings
    df_cfg = xl.parse(0, header=None).fillna('').astype(str).apply(lambda c: c.str.strip().str.lower())
    cfg = dict(zip(df_cfg[0], df_cfg[1]))
    
    tgt = float(cfg.get('threshold', 60))
    w = {k: float(v) for k, v in cfg.items() if any(x in k for x in ('internal','external','assign')) and v.replace('.','').isdigit()} or {'internal': 0.4, 'external': 0.6}
    lvls = sorted([(int(float(k)), float(v)) for k, v in cfg.items() if k.replace('.','').isdigit() and v.replace('.','').isdigit()], reverse=True) or [(3,0.8), (2,0.7), (1,0.6)]

    res, students = {'internal':{}, 'external':{}, 'assign':{}}, {}
    
    # 2. Vectorized Processing per Exam
    for map_sheet in [s for s in xl.sheet_names if 'Mapping' in s]:
        exam = map_sheet.replace(' Ques Mapping', '').replace(' Mapping', '').strip()
        res_sheet = next((s for s in xl.sheet_names if s.startswith(exam) and 'Result' in s), None)
        if not res_sheet: continue
        
        # Mappings
        mdf = xl.parse(map_sheet).rename(columns=lambda x: str(x).strip().upper())
        q_col, m_col = mdf.columns[:2]
        mdf[m_col] = pd.to_numeric(mdf[m_col], errors='coerce').fillna(0)
        
        # CO definitions (robust boolean handling)
        cos = mdf.filter(regex=r'^CO\d+').apply(lambda c: c.apply(lambda v: 1 if v in (1, '1', True, 'True') else 0))
        mdf = mdf[(mdf[m_col] > 0) & (cos.sum(axis=1) > 0)]
        if mdf.empty: continue
        
        # Results
        rdf = xl.parse(res_sheet).rename(columns=lambda x: str(x).strip().upper())
        roll_c = rdf.filter(regex='ROLL|ADMISSION').columns[0]
        name_c = rdf.filter(regex='NAME').columns[0] if len(rdf.filter(regex='NAME').columns) else None
        
        rdf = rdf.dropna(subset=[roll_c]).set_index(roll_c)
        students.update(rdf[name_c].to_dict() if name_c else {r:'' for r in rdf.index})
        marks = rdf.replace(['U','','AB'], np.nan).apply(pd.to_numeric, errors='coerce')
        
        cat = 'external' if 'ETE' in exam.upper() else ('assign' if 'ASN' in exam.upper() else 'internal')
        
        # Fast multi-CO array division
        for co in cos.columns:
            v_qs = mdf.loc[cos[co] > 0, q_col]
            v_qs = v_qs[v_qs.isin(marks.columns)]
            if v_qs.empty: continue
            
            obt = marks[v_qs].sum(axis=1, min_count=1) # NaNs ignored, all-NaN yields NaN
            mx = marks[v_qs].notna().dot(mdf.set_index(q_col).loc[v_qs, m_col]).astype(float)
            res[cat].setdefault(co, []).append(pd.Series(np.where(mx > 0, obt / mx * 100, np.nan), index=obt.index))

    # 3. Aggregation & Multi-weighting
    if not students: raise ValueError("No valid student data found.")
    master = pd.DataFrame({'Name': pd.Series(students)}, index=pd.Index(students.keys(), name='Roll'))
    
    for co in sorted({co for c in res.values() for co in c}):
        w_sum = pd.Series(0.0, index=master.index)
        for cat, weight in w.items():
            k = next((c for c in res if cat[:3] in c), None) # map 'int' to 'internal'
            if k and co in res[k]:
                w_sum += pd.concat(res[k][co], axis=1).mean(axis=1).reindex(master.index).fillna(0) * weight
                
        master[f'{co} %'] = w_sum.replace(0, np.nan).round(2)
        master[f'{co} Target'] = np.where(master[f'{co} %'].isna(), 'N/A', np.where(master[f'{co} %'] >= tgt, 'Yes', 'No'))

    # 4. Success Rates & Export
    sr = master.filter(like='%').apply(lambda x: pd.Series([x.count(), (x >= tgt).sum(), ((x >= tgt).sum() / x.count()) if x.count() else 0], index=['Attempted', 'Met Target', 'Success Rate %'])).T
    sr['Level'] = sr['Success Rate %'].apply(lambda r: next((lv for lv, th in lvls if r >= th), 0))
    sr['Success Rate %'] = (sr['Success Rate %'] * 100).round(2)
    sr.index = [c.split(' ')[0] for c in sr.index]

    with pd.ExcelWriter(out) as ex:
        master.reset_index().to_excel(ex, sheet_name='Student Details', index=False)
        sr.rename_axis('CO').reset_index().to_excel(ex, sheet_name='Summary', index=False)
    return out

# 5. Minimal GUI
def run_app():
    def process_gui():
        f = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if f:
            try: messagebox.showinfo("Done", f"Success! Saved:\n{calc_co(f)}")
            except Exception as e: messagebox.showerror("Error", f"Failed:\n{e}")
            
    r = tk.Tk(); r.title("CO Calc (Quant)"); r.geometry("250x120"); r.eval('tk::PlaceWindow . center')
    tk.Button(r, text="Select Excel File & Run", font=('Arial', 11, 'bold'), bg='#2C3E50', fg='white', command=process_gui).pack(expand=True, fill='both', padx=15, pady=15)
    r.mainloop()

if __name__ == '__main__': run_app()
