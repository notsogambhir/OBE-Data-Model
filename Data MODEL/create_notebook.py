import json

notebook_content = {
 "cells": [
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "# CO Attainment Calculator\n",
    "This notebook calculates Course Outcome (CO) attainments based on student marks, following NBA OBE conventions. It dynamically adapts to the number of COs and reads target thresholds and weightages directly from your input file."
   ]
  },
  {
   "cell_type": "code",
   "execution_count": None,
   "metadata": {},
   "source": [
    "!pip install pandas numpy openpyxl xlrd"
   ],
   "outputs": []
  },
  {
   "cell_type": "code",
   "execution_count": None,
   "metadata": {},
   "source": [
    "import pandas as pd\n",
    "import numpy as np\n",
    "import warnings\n",
    "warnings.filterwarnings('ignore')"
   ],
   "outputs": []
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "### 1. Load Data\n",
    "Please ensure your input file path is correct. If running in Colab, you will need to upload your `Input 1.xlsx` file first. The script will automatically detect the number of exams and COs present in the file."
   ]
  },
  {
   "cell_type": "code",
   "execution_count": None,
   "metadata": {},
   "source": [
    "file_path = \"Input 1.xlsx\" # Update this path if necessary\n",
    "\n",
    "xls = pd.ExcelFile(file_path)\n",
    "sheet_names = xls.sheet_names\n",
    "print(f\"Found sheets: {sheet_names}\")"
   ],
   "outputs": []
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "### 2. Parse Dynamic Variables (Thresholds & Weights)\n",
    "The OBE details are expected to be on the first sheet of the active workbook."
   ]
  },
  {
   "cell_type": "code",
   "execution_count": None,
   "metadata": {},
   "source": [
    "# Assuming the first sheet is the Instructions/OBE Details sheet\n",
    "meta_df = pd.read_excel(xls, sheet_name=sheet_names[0], header=None)\n",
    "\n",
    "# Default values just in case\n",
    "target_percentage = 60.0\n",
    "internal_weight = 0.4\n",
    "external_weight = 0.6\n",
    "level_3_threshold = 0.8\n",
    "level_2_threshold = 0.7\n",
    "level_1_threshold = 0.6\n",
    "\n",
    "try:\n",
    "    target_percentage = float(meta_df[meta_df[0] == 'Threshold'].iloc[0, 1])\n",
    "except: pass\n",
    "\n",
    "try:\n",
    "    internal_weight = float(meta_df[meta_df[0].astype(str).str.contains('Internal', na=False, case=False)].iloc[0, 1])\n",
    "    external_weight = float(meta_df[meta_df[0].astype(str).str.contains('External', na=False, case=False)].iloc[0, 1])\n",
    "except: pass\n",
    "\n",
    "try:\n",
    "    level_3_threshold = float(meta_df[meta_df[0] == 3].iloc[0, 1])\n",
    "    level_2_threshold = float(meta_df[meta_df[0] == 2].iloc[0, 1])\n",
    "    level_1_threshold = float(meta_df[meta_df[0] == 1].iloc[0, 1])\n",
    "except: pass\n",
    "\n",
    "print(f\"Target Course Percentage: {target_percentage}%\")\n",
    "print(f\"Internal Weight: {internal_weight} | External Weight: {external_weight}\")\n",
    "print(f\"Levels Class Attainment Targets -> L3: {level_3_threshold*100}%, L2: {level_2_threshold*100}%, L1: {level_1_threshold*100}%\")"
   ],
   "outputs": []
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "### 3. Discover Exams and COs\n",
    "Dynamically scans the file to figure out what exams exist (ST1, ST2, ETE etc.) and extracts all unique COs across all mapping sheets."
   ]
  },
  {
   "cell_type": "code",
   "execution_count": None,
   "metadata": {},
   "source": [
    "exam_names = []\n",
    "for sheet in sheet_names:\n",
    "    if 'Ques Mapping' in sheet or 'Mapping' in sheet:\n",
    "        exam_names.append(sheet.replace(' Ques Mapping', '').replace(' Mapping', ''))\n",
    "        \n",
    "print(f\"Detected Exams: {exam_names}\")\n",
    "\n",
    "all_mapping_dfs = {}\n",
    "all_result_dfs = {}\n",
    "all_cos = set()\n",
    "\n",
    "for exam in exam_names:\n",
    "    # Try to find corresponding sheets\n",
    "    map_sheet = next((s for s in sheet_names if exam in s and 'Mapping' in s), None)\n",
    "    res_sheet = next((s for s in sheet_names if exam in s and 'Result' in s), None)\n",
    "    \n",
    "    if map_sheet and res_sheet:\n",
    "        map_df = pd.read_excel(xls, sheet_name=map_sheet)\n",
    "        res_df = pd.read_excel(xls, sheet_name=res_sheet)\n",
    "        \n",
    "        all_mapping_dfs[exam] = map_df\n",
    "        all_result_dfs[exam] = res_df\n",
    "        \n",
    "        co_cols = [str(c) for c in map_df.columns if str(c).upper().startswith('CO')]\n",
    "        all_cos.update(co_cols)\n",
    "\n",
    "all_cos = sorted(list(all_cos))\n",
    "print(f\"Detected COs: {all_cos}\")"
   ],
   "outputs": []
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "### 4. Calculation Engine\n",
    "The function below aggregates marks exactly as described in the requirements: only attempted questions are factored into the maximum marks denominator."
   ]
  },
  {
   "cell_type": "code",
   "execution_count": None,
   "metadata": {},
   "source": [
    "def process_exam_for_student(student_row, map_df, co_name):\n",
    "    if co_name not in map_df.columns:\n",
    "        return np.nan\n",
    "    \n",
    "    # Convert CO mapping to numeric cleanly\n",
    "    co_mapping = pd.to_numeric(map_df[co_name], errors='coerce').fillna(0)\n",
    "    valid_questions = map_df[co_mapping > 0]\n",
    "    \n",
    "    if valid_questions.empty:\n",
    "        return np.nan\n",
    "    \n",
    "    obtained_total = 0\n",
    "    max_total = 0\n",
    "    \n",
    "    for idx, row in valid_questions.iterrows():\n",
    "        # Use column Q_Id or infer the question column name\n",
    "        q_col = None\n",
    "        if 'Q_Id' in row:\n",
    "            q_col = str(row['Q_Id']).strip()\n",
    "        else:\n",
    "            q_col = row.iloc[0] # assume first col is question id\n",
    "            \n",
    "        max_marks = pd.to_numeric(row.get('Max Marks', np.nan), errors='coerce')\n",
    "        if pd.isna(max_marks):\n",
    "            max_marks = pd.to_numeric(row.iloc[1], errors='coerce')\n",
    "            \n",
    "        if pd.isna(max_marks):\n",
    "            continue\n",
    "            \n",
    "        if q_col in student_row.index:\n",
    "            marks_obtained = student_row[q_col]\n",
    "            marks_num = pd.to_numeric(marks_obtained, errors='coerce')\n",
    "            \n",
    "            if pd.notna(marks_num): # Student Attempted\n",
    "                obtained_total += marks_num\n",
    "                max_total += max_marks\n",
    "                \n",
    "    if max_total == 0:\n",
    "        return np.nan\n",
    "        \n",
    "    return (obtained_total / max_total) * 100\n"
   ],
   "outputs": []
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "### 5. Executing Calculations For All Students\n",
    "Loops over all unique students, exams and COs to compile the overall attainment statistics."
   ]
  },
  {
   "cell_type": "code",
   "execution_count": None,
   "metadata": {},
   "source": [
    "# Get unique list of students from all result sheets\n",
    "all_students = []\n",
    "for res_df in all_result_dfs.values():\n",
    "    if 'Admission No. (Roll No.)' in res_df.columns:\n",
    "        sub_df = res_df[['Admission No. (Roll No.)', 'Name of Student']].copy().dropna(subset=['Admission No. (Roll No.)'])\n",
    "        all_students.append(sub_df)\n",
    "        \n",
    "students_df = pd.concat(all_students).drop_duplicates(subset=['Admission No. (Roll No.)']).set_index('Admission No. (Roll No.)')\n",
    "\n",
    "final_attainment = {}\n",
    "student_co_master = students_df.copy()\n",
    "\n",
    "for co in all_cos:\n",
    "    co_data = pd.DataFrame(index=students_df.index)\n",
    "    \n",
    "    internal_exams = [e for e in exam_names if 'ETE' not in e.upper() and 'EXTERNAL' not in e.upper()]\n",
    "    external_exams = [e for e in exam_names if 'ETE' in e.upper() or 'EXTERNAL' in e.upper()]\n",
    "\n",
    "    # Compute Internal\n",
    "    for exam in internal_exams:\n",
    "        map_df = all_mapping_dfs[exam]\n",
    "        res_df = all_result_dfs[exam].set_index('Admission No. (Roll No.)')\n",
    "        scores = [process_exam_for_student(res_df.loc[sid], map_df, co) if sid in res_df.index else np.nan for sid in students_df.index]\n",
    "        co_data[f\"{exam}_%\"] = scores\n",
    "        \n",
    "    if internal_exams:\n",
    "        co_data['Internal_Avg'] = co_data[[f\"{e}_%\" for e in internal_exams]].mean(axis=1, skipna=True)\n",
    "    else:\n",
    "        co_data['Internal_Avg'] = np.nan\n",
    "        \n",
    "    # Compute External\n",
    "    for exam in external_exams:\n",
    "        map_df = all_mapping_dfs[exam]\n",
    "        res_df = all_result_dfs[exam].set_index('Admission No. (Roll No.)')\n",
    "        scores = [process_exam_for_student(res_df.loc[sid], map_df, co) if sid in res_df.index else np.nan for sid in students_df.index]\n",
    "        co_data[f\"{exam}_%\"] = scores\n",
    "        \n",
    "    ext_col = f\"{external_exams[0]}_%\" if external_exams else None\n",
    "    \n",
    "    # Apply Weighted Totals\n",
    "    if ext_col and ext_col in co_data.columns and 'Internal_Avg' in co_data.columns:\n",
    "        # Handle cases where a student missed internal or external\n",
    "        int_val = co_data['Internal_Avg'].fillna(0)\n",
    "        ext_val = co_data[ext_col].fillna(0)\n",
    "        # If both are NaN, Total is NaN. If one is NaN, it's 0 (effectively penalized) or depends on policy. We assume penalty as 0.\n",
    "        co_data['Total_%'] = np.where(co_data['Internal_Avg'].isna() & co_data[ext_col].isna(), np.nan, (int_val * internal_weight) + (ext_val * external_weight))\n",
    "    elif 'Internal_Avg' in co_data.columns:\n",
    "        co_data['Total_%'] = co_data['Internal_Avg']\n",
    "    elif ext_col:\n",
    "        co_data['Total_%'] = co_data[ext_col]\n",
    "    else:\n",
    "        co_data['Total_%'] = np.nan\n",
    "        \n",
    "    # Check attainment target\n",
    "    co_data['Attained Target'] = co_data['Total_%'] >= target_percentage\n",
    "    \n",
    "    # Compute success rate (excluding students who were entirely absent for this CO)\n",
    "    valid_students = co_data['Total_%'].notna()\n",
    "    success_count = co_data.loc[valid_students, 'Attained Target'].sum()\n",
    "    total_count = valid_students.sum()\n",
    "    success_rate = (success_count / total_count) if total_count > 0 else 0\n",
    "    \n",
    "    # Assign Class Level\n",
    "    if success_rate >= level_3_threshold:\n",
    "        level = 3\n",
    "    elif success_rate >= level_2_threshold:\n",
    "        level = 2\n",
    "    elif success_rate >= level_1_threshold:\n",
    "        level = 1\n",
    "    else:\n",
    "        level = 0\n",
    "        \n",
    "    final_attainment[co] = {\n",
    "        'Students Attempted': total_count,\n",
    "        'Students Passed target': success_count,\n",
    "        'Success Rate (%)': round(success_rate * 100, 2),\n",
    "        'Attainment Level': level\n",
    "    }\n",
    "    \n",
    "    # Save their total back to master dataframe for export\n",
    "    student_co_master[f\"{co} %\"] = co_data['Total_%'].round(2)\n",
    "\n",
    "attainment_df = pd.DataFrame(final_attainment).T\n",
    "attainment_df.index.name = 'Course Outcome'\n",
    "display(attainment_df)"
   ],
   "outputs": []
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "### 6. Export Results\n",
    "The summary mapping alongside individual student percentages is saved to a newly generated Excel file."
   ]
  },
  {
   "cell_type": "code",
   "execution_count": None,
   "metadata": {},
   "source": [
    "output_file = \"Output_Attainment.xlsx\"\n",
    "with pd.ExcelWriter(output_file) as writer:\n",
    "    attainment_df.to_excel(writer, sheet_name=\"Class Attainment Summary\")\n",
    "    student_co_master.to_excel(writer, sheet_name=\"Student Details\")\n",
    "\n",
    "print(f\"Processing complete! Results saved to {output_file}\")"
   ],
   "outputs": []
  }
 ],
 "metadata": {
  "kernelspec": {
   "display_name": "Python 3",
   "language": "python",
   "name": "python3"
  },
  "language_info": {
   "codemirror_mode": {
    "name": "ipython",
    "version": 3
   },
   "file_extension": ".py",
   "mimetype": "text/x-python",
   "name": "python",
   "nbconvert_exporter": "python",
   "pygments_lexer": "ipython3",
   "version": "3.10.12"
  }
 },
 "nbformat": 4,
 "nbformat_minor": 4
}

with open(r'd:\Data MODEL\CO_Attainment_Calculator.ipynb', 'w', encoding='utf-8') as f:
    json.dump(notebook_content, f, indent=1)
    
print("Notebook created continuously!")
