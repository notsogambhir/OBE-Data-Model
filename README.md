# NBA OBE Data Model & Automation Project

Welcome to the central repository for Outcome-Based Education (OBE) data management and CO Attainment tracking. This project embodies the evolution of our CO Attainment calculation methodology, moving from a fully manual Excel-based framework (Phase 1) to a decoupled, automated Python-driven engine (Phase 2).

This document serves as the global guide explaining the end-to-end workflows of both phases.

---

## 🏗️ Phase 1: Excel-Based OBE Management

**Directory:** `/Excel/`

In the initial phase of our OBE implementation, the entire ecosystem—from data collection to complex aggregations—was managed entirely within Microsoft Excel (`Combined_legacy/` and `Individual_Updated/`). 

### End-to-End Workflow (Phase 1)

1. **Workbook Structuring:** A massive Excel workbook was created for a specific course merging multiple assessments (Quizzes, Mid-Terms, End-Terms).
2. **Manual CO Mapping:** Teachers would set up a matrix at the top of the spreadsheet to strictly map each question column to specific Course Outcomes (CO1, CO2, etc.).
3. **Raw Data Entry:** Instructors manually entered raw student marks for each question across hundreds of rows.
4. **Individual Calculation (Formulas):** Complex `SUM(IF(...))` or `SUMPRODUCT` formulas were built at the end of each student row to:
   - Sum the marks the student earned for questions mapped to a specific CO.
   - Sum the maximum marks possible for that specific CO.
   - Calculate the percentage `(Obtained / Max) * 100`.
5. **Target Thresholding:** Nested `IF` statements evaluated whether the calculated percentage met the predefined department target (e.g., `>= 60%`), marking the cell with a `1` (Attained) or `0` (Not Attained).
6. **Class Aggregation:** At the bottom of the worksheet, `COUNTIF` formulas aggregated the successes to determine the **Class Success Rate** (e.g., 65% of the class attained CO1).
7. **Level Assignment:** Further nested `IF` statements converted the success rate into an Attainment Level (e.g., Level 1, 2, or 3).
8. **Cross-Exam Integration:** Final CO attainment across the entire semester required manually linking cells from different interconnected workbook sheets to apply Internal (e.g., 40%) vs. External (e.g., 60%) weightages.

### Limitations Addressed by Phase 2
- **Hardcoded & Fragile:** Changing the number of COs (e.g., from 4 to 6) or adding questions required dragging and modifying hundreds of complex formulas.
- **Unattempted Question Errors:** Excel formulas struggled to fairly drop unattempted questions from the denominator without creating massive, computationally heavy array formulas.
- **High Friction:** extremely prone to human error through broken cell references.

---

## 🚀 Phase 2: Python-Based Data Model

**Directory:** `/Data MODEL/`

Phase 2 acts as a paradigm shift. We decouple the raw data entry (Excel) from the calculation logic (Python). Excel is now used strictly as a "dumb" data intake template, while Python scripts handle all filtering, weighting, and tier-based aggregations securely in the background.

### End-to-End Workflow (Phase 2)

#### Step 1: Data Ingestion
1. **Template Preparation:** The user populates standardized Excel templates (`Input 1.xlsx`, etc.). These files have a fixed, simple structure:
   - **`OBE Details` Sheet:** The configuration hub. The user defines global variables here: Target Percentage (e.g., 60%), Assessment Weights (e.g., Internal 40%, External 60%), and Attainment Level thresholds (e.g., L3 >= 70%).
   - **`[Exam] Mapping` Sheets:** A simple table mapping each question ID to its Max Marks and its relevance (1 or 0) to specific COs.
   - **`[Exam] Result` Sheets:** Pure raw student mark entries.

#### Step 2: Engine Execution
2. The user runs the calculation engine. Depending on preference, they can run:
   - `GEM_app.py`: A user-friendly desktop GUI where the user browses for the input file and clicks "Run".
   - `GEM_fixed.py` / `co_attainment_pandas.py`: Pandas-accelerated calculation engines for CLI processing.
   - `co_attainment.py`: A lightweight standard-library python alternative.

#### Step 3: Dynamic Calculation (Tier 1 - Individual)
3. **Parsing & Mapping:** The script dynamically parses the `OBE Details`. It doesn't matter if there are 2 COs or 12 COs; the script adapts automatically by reading the mapping headers.
4. **Fair Grading Algorithm:** The engine iterates through every student's record. For every mapped question, it calculates the CO percentage. **Crucially, if a student left a question blank (unattempted/absent), the script dynamically removes that specific question's max marks from their individual denominator**, ensuring fair evaluation.

#### Step 4: Weighted Integration (Tier 2 - Cross Assessment)
5. **Category Grouping:** The engine groups all exams (e.g., grouping `ST1` and `ST2` into the `Internal` category based on naming conventions).
6. **Averaging & Weighting:** It averages the CO percentages within each category, and then multiplies them by the weights defined in the `OBE Details` sheet to compute a single **Final Weighted CO Percentage** per student.

#### Step 5: Class Evaluation (Tier 3 - Global)
7. **Success Counting:** The engine evaluates every student's Final Weighted CO Percentage against the Target Threshold to compute the overall **Class Success Rate** for each CO.
8. **Level Mapping:** The computed Success Rate is compared against the Attainment Levels thresholds to lock in the final Level (0, 1, 2, or 3).

#### Step 6: Output Generation
9. **Automated Reporting:** The Python engine generates a pristine `Output_[Filename].xlsx` without altering the original input file. The output contains:
   - **Student Details:** A complete breakdown of every student's CO percentages and whether they met the target.
   - **Attainment Summary:** The high-level analytics table defining the final Attainment levels across the class.
   - **Configuration:** A stamp of the weights and targets used for the calculation, ensuring historical auditability.

---

## 🚀 How to Utilize This Repository
1. For historical data or strictly manual template structures, reference the `/Excel/` folder.
2. For processing new assessment data, fill out the standardized `.xlsx` templates and run them through the apps provided in the `/Data MODEL/` directory for automated processing.

*Designed to simplify, automate, and standardize NBA OBE Compliance across all courses.*
