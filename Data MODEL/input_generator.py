import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import os

class InputGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Input File Generator for CO Attainment")
        self.root.geometry("600x750")
        self.root.resizable(False, False)
        
        # Data structure to hold exam details
        self.exams = []
        
        self.setup_ui()
        
    def setup_ui(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        lbl_title = tk.Label(main_frame, text="Generate Custom Input Excel File", font=("Arial", 14, "bold"))
        lbl_title.pack(pady=(0, 10))
        
        # ─── OBE CONFIGURATION ───
        frame_obe = tk.LabelFrame(main_frame, text="OBE Configuration", padx=10, pady=10, font=("Arial", 10, "bold"))
        frame_obe.pack(fill=tk.X, pady=5)
        
        tk.Label(frame_obe, text="Target Threshold (%):").grid(row=0, column=0, sticky='w', pady=5)
        self.threshold_var = tk.DoubleVar(value=60.0)
        tk.Entry(frame_obe, textvariable=self.threshold_var, width=15).grid(row=0, column=1, padx=10, pady=5)
        
        tk.Label(frame_obe, text="Level 3 Target (%):").grid(row=1, column=0, sticky='w', pady=5)
        self.level3_var = tk.DoubleVar(value=80.0)
        tk.Entry(frame_obe, textvariable=self.level3_var, width=15).grid(row=1, column=1, padx=10, pady=5)
        
        tk.Label(frame_obe, text="Level 2 Target (%):").grid(row=2, column=0, sticky='w', pady=5)
        self.level2_var = tk.DoubleVar(value=70.0)
        tk.Entry(frame_obe, textvariable=self.level2_var, width=15).grid(row=2, column=1, padx=10, pady=5)
        
        tk.Label(frame_obe, text="Level 1 Target (%):").grid(row=3, column=0, sticky='w', pady=5)
        self.level1_var = tk.DoubleVar(value=60.0)
        tk.Entry(frame_obe, textvariable=self.level1_var, width=15).grid(row=3, column=1, padx=10, pady=5)
        
        # ─── ADD EXAM CONFIGURATION ───
        frame_top = tk.LabelFrame(main_frame, text="Add Exam Configuration", padx=10, pady=10, font=("Arial", 10, "bold"))
        frame_top.pack(fill=tk.X, pady=5)
        
        # Use a grid for the input fields
        tk.Label(frame_top, text="Exam Name (e.g. ST1, ETE):").grid(row=0, column=0, sticky='w', pady=5)
        self.exam_name_var = tk.StringVar()
        tk.Entry(frame_top, textvariable=self.exam_name_var, width=25).grid(row=0, column=1, padx=10, pady=5)
        
        tk.Label(frame_top, text="Number of COs:").grid(row=1, column=0, sticky='w', pady=5)
        self.co_var = tk.IntVar(value=6)
        tk.Entry(frame_top, textvariable=self.co_var, width=25).grid(row=1, column=1, padx=10, pady=5)
        
        tk.Label(frame_top, text="Number of Questions:").grid(row=2, column=0, sticky='w', pady=5)
        self.q_var = tk.IntVar(value=10)
        tk.Entry(frame_top, textvariable=self.q_var, width=25).grid(row=2, column=1, padx=10, pady=5)
        
        btn_add = tk.Button(frame_top, text="Add Exam", command=self.add_exam, bg="#2196F3", fg="white", font=("Arial", 10, "bold"), width=15)
        btn_add.grid(row=3, column=0, columnspan=2, pady=10)
        
        # ─── EXAMS LIST ───
        frame_mid = tk.LabelFrame(main_frame, text="Exams to be added", padx=10, pady=10, font=("Arial", 10, "bold"))
        frame_mid.pack(fill=tk.BOTH, expand=True, pady=5)
        
        columns = ("Exam Name", "COs", "Questions")
        self.tree = ttk.Treeview(frame_mid, columns=columns, show='headings', height=6)
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150, anchor=tk.CENTER)
        self.tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        
        scrollbar = ttk.Scrollbar(frame_mid, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        btn_remove = tk.Button(main_frame, text="Remove Selected Exam", command=self.remove_exam, width=20)
        btn_remove.pack(pady=5)
        
        # ─── GENERATE BUTTON ───
        frame_bottom = tk.Frame(main_frame, pady=10)
        frame_bottom.pack(fill=tk.X)
        
        btn_gen = tk.Button(frame_bottom, text="Generate Input Excel File...", command=self.generate_file, 
                  bg="#4CAF50", fg="white", font=("Arial", 12, "bold"), height=2)
        btn_gen.pack(fill=tk.X)
                  
    def add_exam(self):
        name = self.exam_name_var.get().strip()
        try:
            cos = self.co_var.get()
            qs = self.q_var.get()
        except tk.TclError:
            messagebox.showerror("Error", "COs and Questions must be valid integers.")
            return
            
        if not name:
            messagebox.showerror("Error", "Exam name cannot be empty.")
            return
        if cos <= 0 or qs <= 0:
            messagebox.showerror("Error", "COs and Questions must be greater than 0.")
            return
            
        # Check if exam name already exists
        for exam in self.exams:
            if exam["name"].lower() == name.lower():
                messagebox.showerror("Error", f"Exam '{name}' has already been added.")
                return
                
        self.exams.append({"name": name, "cos": cos, "qs": qs})
        self.tree.insert("", tk.END, values=(name, cos, qs))
        self.exam_name_var.set("")
        
    def remove_exam(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select an exam to remove.")
            return
        for item in selected:
            idx = self.tree.index(item)
            self.exams.pop(idx)
            self.tree.delete(item)
            
    def generate_file(self):
        if not self.exams:
            messagebox.showerror("Error", "Please add at least one exam configuration.")
            return
            
        try:
            threshold = self.threshold_var.get()
            l3 = self.level3_var.get() / 100.0
            l2 = self.level2_var.get() / 100.0
            l1 = self.level1_var.get() / 100.0
        except tk.TclError:
            messagebox.showerror("Error", "OBE Configuration targets must be valid numbers.")
            return
            
        path = filedialog.asksaveasfilename(
            title="Save Input Excel File As",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="Custom_Input.xlsx"
        )
        if not path:
            return
            
        try:
            # Create a Pandas Excel writer using openpyxl as the engine
            with pd.ExcelWriter(path, engine='openpyxl') as writer:
                # 1. Write OBE Details sheet
                obe_data = {
                    "CO Score": ["Threshold", None, None, "Types", "Internal (Avg of ST1,ST2,ST3)", "External(ETE)", None, "CO Score", 3, 2, 1],
                    "Percentage of Co attained": [threshold, None, None, "Weightages", 0.4, 0.6, None, "Perecntage of students attaining Target", l3, l2, l1]
                }
                df_obe = pd.DataFrame(obe_data)
                df_obe.to_excel(writer, sheet_name="OBE Details", index=False, header=False)
                
                # 2. For each exam, create '{Exam} Ques Mapping' and '{Exam} Result' sheets
                for exam in self.exams:
                    name = exam["name"]
                    cos = exam["cos"]
                    qs = exam["qs"]
                    
                    q_ids = [f"Q{i+1}" for i in range(qs)]
                    
                    # --- Ques Mapping Sheet ---
                    mapping_cols = ["Q_Id", "Max Marks"] + [f"CO{i+1}" for i in range(cos)]
                    df_map = pd.DataFrame(columns=mapping_cols)
                    df_map["Q_Id"] = q_ids
                    df_map["Max Marks"] = 1 # Default max marks
                    
                    for i in range(cos):
                        df_map[f"CO{i+1}"] = False
                    
                    df_map.to_excel(writer, sheet_name=f"{name} Ques Mapping", index=False)
                    
                    # --- Result Sheet ---
                    res_cols = ["Sr.no", "Admission No. (Roll No.)", "Name of Student", "Course Code", f"Exam Name ({name})"] + q_ids + ["Total Marks", "max Marks"]
                    # Add some empty rows to make it clear where to enter data
                    df_res = pd.DataFrame([[None]*len(res_cols)]*5, columns=res_cols)
                    
                    # Pre-fill Sr.no for convenience
                    df_res["Sr.no"] = range(1, 6)
                    
                    df_res.to_excel(writer, sheet_name=f"{name} Result", index=False)
                    
            messagebox.showinfo("Success", f"Input file successfully generated at:\n{path}\n\nYou can now fill it with your data and use it in GEM_app.py.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = InputGeneratorApp(root)
    root.mainloop()
