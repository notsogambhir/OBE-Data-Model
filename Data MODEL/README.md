# 📊 CO Attainment Data Model

A zero-dependency Python tool that calculates **Course Outcome (CO) Attainment Levels** from student marks, following the **NBA OBE (Outcome-Based Education)** framework.

---

## 📁 Project Structure

```
Data MODEL/
├── co_attainment.py                 # Standalone calculation script
├── CO_Attainment_Calculator.ipynb   # Jupyter/Colab notebook (detailed walkthrough)
├── Input 1.xlsx                     # Sample input (6 COs, 2-weight)
├── Input 2.xlsx                     # Sample input (6 COs, different COs per exam)
├── Input 3.xlsx                     # Sample input (6 COs, 3-weight with Assignment)
├── Output_Input 1.xlsx              # Generated output
├── Output_Input 2.xlsx              # Generated output
├── Output_Input 3.xlsx              # Generated output
├── nb_part1.py                      # Notebook generator (Part 1)
├── nb_part2.py                      # Notebook generator (Part 2)
└── README.md                        # This file
```

---

## 🚀 Quick Start

### Option 1: Python Script
```bash
python co_attainment.py "Input 1.xlsx" "Input 2.xlsx" "Input 3.xlsx"
```

### Option 2: Jupyter / Colab Notebook
1. Open `CO_Attainment_Calculator.ipynb`
2. Set `INPUT_FILES` in Section 2
3. **Run All** cells

> **No external libraries required.** The tool uses only Python's standard library (`zipfile`, `xml`).

---

## 📐 Data Model Overview

### Input File Structure

Each input `.xlsx` file must contain these sheet types:

| Sheet | Purpose | Example |
|---|---|---|
| **OBE Details** | Configuration (target, weights, levels) | Always the **first** sheet |
| **\<Exam\> Ques Mapping** | Question → CO mapping matrix | `ST1 Ques Mapping` |
| **\<Exam\> Result** | Student marks per question | `ST1 Result` |

### 1. OBE Details Sheet (Configuration)

This sheet defines all calculation parameters:

```
┌──────────────────────────────────┬────────────┐
│ CO Score                         │ Percentage │
├──────────────────────────────────┼────────────┤
│ Threshold                        │ 60         │  ← Target %
│                                  │            │
│ Types                            │ Weightages │
│ Internal (Avg of ST1,ST2,ST3)    │ 0.4        │  ← Weight categories
│ External(ETE)                    │ 0.6        │
│ Assignment                       │ 0.1        │  ← Optional
│                                  │            │
│ CO Score                         │ % students │
│ 3                                │ 0.8        │  ← Level thresholds
│ 2                                │ 0.7        │
│ 1                                │ 0.6        │
└──────────────────────────────────┴────────────┘
```

**Dynamic features:**
- Supports **2-weight** (Internal + External) or **3-weight** (+ Assignment) schemes
- All thresholds and targets are read from the sheet, not hardcoded

### 2. Question Mapping Sheet

Defines which questions test which COs (many-to-many relationship):

```
┌───────┬───────────┬─────┬─────┬─────┬─────┐
│ Q_Id  │ Max Marks │ CO1 │ CO2 │ CO3 │ CO4 │
├───────┼───────────┼─────┼─────┼─────┼─────┤
│ Q1    │ 5         │  1  │  0  │  0  │  1  │  ← Q1 tests CO1 AND CO4
│ Q2    │ 2         │  0  │  1  │  0  │  0  │  ← Q2 tests only CO2
│ Q3    │ 10        │  1  │  1  │  0  │  0  │  ← Q3 tests CO1 AND CO2
└───────┴───────────┴─────┴─────┴─────┴─────┘
```

**Dynamic features:**
- Handles **2 to 8+ COs** — detected from column headers
- Different exams can have **different CO columns** (e.g., ST1 has CO1/CO2, ST2 has CO3/CO4)
- Questions with max marks = 0 are automatically skipped (handles empty ST3 sheets)

### 3. Result Sheet

Contains student marks per question:

```
┌───────┬─────────────┬────────┬────────┬─────┬────┬────┬────┬───────┐
│ Sr.no │ Admission   │ Name   │ Course │ Exam│ Q1 │ Q2 │ Q3 │ Total │
├───────┼─────────────┼────────┼────────┼─────┼────┼────┼────┼───────┤
│ 1     │ 2010990089  │ ANSH   │ EC114  │ ST1 │ 4  │ U  │ 8  │ 12    │
│ 2     │ 2010992002  │ PANSY  │ EC114  │ ST1 │ 5  │ 2  │    │ 7     │
└───────┴─────────────┴────────┴────────┴─────┴────┴────┴────┴───────┘
```

**Unattempted markers** (all treated identically):
- `U` — explicitly marked unattempted
- *(blank/empty)* — no entry
- `AB` — absent

---

## 🧮 Calculation Pipeline

```
Step 1                Step 2                Step 3              Step 4
Parse OBE     →    Discover Exams   →   Parse Mappings   →  Per-Student
Config               (ST1, ETE...)       & Results           CO% per Exam
                                                                 │
                                                                 ▼
Step 8                Step 7               Step 6              Step 5
Assign Level  ←   Class Success    ←   Target Check     ←  Weighted
(0, 1, 2, 3)       Rate per CO         per Student         Average
```

### Step-by-Step Detail

#### Step 4: Per-Student CO Percentage (Core Formula)

```
              Σ (obtained marks on attempted questions mapped to this CO)
CO % = ──────────────────────────────────────────────────────────────────── × 100
              Σ (max marks on attempted questions mapped to this CO)
```

**Critical rule:** Unattempted questions are excluded from **both** numerator and denominator.

| Question | Max | CO1 Mapped | Score | Included? |
|---|---|---|---|---|
| Q1 | 10 | ✅ | 8 | ✅ Yes (attempted + mapped) |
| Q2 | 10 | ✅ | U | ❌ No (unattempted) |
| Q3 | 5 | ❌ | 4 | ❌ No (not mapped to CO1) |

**CO1% = 8 / 10 = 80%** *(only Q1 counts)*

#### Step 5: Weighted Average

```
Final CO% = (Internal_Avg × W_int) + (External% × W_ext) + (Assignment% × W_asn)
```

Where `Internal_Avg` = mean of all ST exam CO percentages for that student.

#### Step 6–8: Class Aggregation

```
Success Rate = (Students meeting target) / (Students with valid data) × 100

Level 3  if  Success Rate ≥ 80%
Level 2  if  Success Rate ≥ 70%
Level 1  if  Success Rate ≥ 60%
Level 0  otherwise
```

---

## 📤 Output Format

Each output `.xlsx` file contains 3 sheets:

| Sheet | Contents |
|---|---|
| **Student Details** | Per-student final CO%, target met (Yes/No) |
| **Attainment Summary** | Per-CO: attempted, met target, success rate, level |
| **Configuration** | Target %, weights, thresholds used |

---

## 🔧 Flexibility & Edge Cases

| Scenario | How It's Handled |
|---|---|
| 2 to 8+ COs per exam | CO columns detected dynamically from headers |
| Variable exams (ST1/ST2/ST3/ETE/ASN) | Discovered automatically from sheet names |
| 2-weight or 3-weight schemes | Read from OBE Details sheet |
| Empty ST3 (all zeros) | Automatically skipped |
| Student absent from one exam | That exam contributes nothing; others still counted |
| All questions unattempted for a CO | Returns N/A (not 0%) |
| Many-to-many Q→CO mapping | Question marks contribute fully to ALL mapped COs |
| Non-breaking spaces in names | Cleaned automatically |

---

## 📖 Reference

Based on:
- **NBA OBE Framework** — National Board of Accreditation, Outcome-Based Education
- [11_Process_CO_Attainment_from_Marks.md](11_Process_CO_Attainment_from_Marks.md) — Detailed process specification
