# CO Attainment from Student Marks

A comprehensive, automated pipeline for calculating Course Outcome (CO) attainment levels based on raw student marks in an Outcome-Based Education (OBE) framework.

## 📌 Overview

In Outcome-Based Education, student success is measured not just by total scores, but by proficiency in specific **Course Outcomes (COs)**. This project provides a calculation engine that processes raw student assessment data, evaluates it against predefined targets, applies weighted aggregations across multiple exam types, and outputs final class-level attainment levels.

This repository houses multiple implementations of the calculation engine, including pandas-based versions (`GEM_fixed.py`, `co_attainment_pandas.py`) and a lightweight, standard-library-only version (`co_attainment.py`).

---

## 🔄 End-to-End Data Flow

The calculation process follows a strict, step-by-step pipeline from raw data ingestion to final level assignment.

### 1. The Input (Excel Files)
The system reads assessment data from standardized Excel files. Each file represents a course and contains three types of sheets:
- **`OBE Details`**: The configuration sheet. Defines the target threshold percentage (e.g., 60%), default weights for different assessment categories (Internal vs. External), and the class success thresholds for Attainment Levels (e.g., Level 3 if 80%+ students meet the target).
- **`[Exam Name] Mapping`**: Defines the assessment structure. Maps individual questions (Q1, Q2, etc.) to one or more COs and specifies the maximum marks for each question.
- **`[Exam Name] Result`**: Contains raw marks obtained by each student on every question.

### 2. Tier 1: Individual Student CO Performance
For every student in a specific exam, the engine calculates their proficiency in each mapped CO:
- **Identify Linked Questions**: Finds all questions mapped to a specific CO.
- **Filter Unattempted**: If a student did not attempt a question (marked as empty, 'U', 'AB', or NaN), it is excluded from both the maximum marks and obtained marks. This ensures students are not penalized for unattempted/optional questions.
- **Calculate Percentage**: `(Sum of Obtained Marks / Sum of Max Marks) * 100`

### 3. Tier 2: Weighted Integration Across Exams
Exams are categorized (Internal, External, Assignment) based on their names (e.g., "ST1", "ETE", "ASN").
- The system averages the CO percentages for all exams within a single category.
- It then applies the weights defined in the `OBE Details` sheet to calculate a **Final Weighted CO Percentage** for each student.

### 4. Tier 3: Class-Level Attainment Evaluation
Once individual weighted scores are calculated, the engine zooms out to the class level:
- **Target Filtering**: For each CO, it counts how many students achieved a score greater than or equal to the **Target Percentage** (e.g., 60%).
- **Success Rate**: Calculates the percentage of the class that met the target (`Students Meeting Target / Total Valid Students`).
- **Level Assignment**: Computes the final Attainment Level (0, 1, 2, or 3) by comparing the Success Rate against the `OBE Details` thresholds.

---

## 🛠️ Technical Implementation & Usage

### Core Scripts
- **`co_attainment.py`**: A robust, dependency-free implementation that uses Python's built-in `zipfile` and `xml.etree` to parse `.xlsx` files. Ideal for environments where `pandas` is unavailable.
- **`GEM_fixed.py` / `co_attainment_pandas.py`**: Pandas-dependent implementations that offer concise vector-based calculations and utilize `openpyxl`.

### How to Run
You can process a single file or multiple files in batch:
```bash
# Process a specific file (Dependencies: pandas, openpyxl for GEM_fixed.py)
python GEM_fixed.py "Input 1.xlsx"

# Using the dependency-free script
python co_attainment.py "Input 1.xlsx"
```
If no file is specified, the script automatically processes all files in the directory matching the `Input*.xlsx` pattern.

### The Output
For each processed input file, the script generates an `Output_[filename].xlsx` file alongside it. The output contains:
1. **Student Details**: A comprehensive sheet showing every student's Roll Number, Name, calculated % for each CO, and a simple Yes/No on whether they met the target.
2. **Attainment Summary**: The high-level overview detailing the total students attempted, number of students meeting the target, overall success rate, and the final evaluated Attainment Level for each CO.
3. **Configuration / Settings**: A reflection of the targets, weights, and thresholds used during calculation for transparency and auditing.

---

## 🛡️ Key Data Integrity Rules Enforced
1. **Dynamic Scaling**: Supports any arbitrary number of COs (CO1 to CO8+) and an unlimited number of questions per exam.
2. **Fair Grading**: Unattempted questions natively bypass the grading denominator for the individual, preventing unfair percentage drops.
3. **Weightage Normalization**: Proper proportional distribution of weights depending on the assessment category logic derived dynamically from input files.
4. **Multiple CO Mapping**: A single question can map to multiple COs; the calculation engine respects shared impact correctly.
