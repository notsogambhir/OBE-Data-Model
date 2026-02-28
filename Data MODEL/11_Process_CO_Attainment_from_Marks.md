# PRD 11: Process of CO Attainment from Marks (In-Depth)

This document provides an exhaustive, step-by-step explanation of how raw student marks are transformed into final Course Outcome (CO) attainment levels within the NBA OBE Portal. It bridges the gap between assessment management and the calculation engine.

---

## 1. The Conceptual Framework

In Outcome-Based Education (OBE), we don't just care about a student's total score in a subject. We care about their proficiency in specific **Course Outcomes (COs)**. 

The process follows this logical chain:
1.  **Assessment Design:** Questions are created to test specific COs.
2.  **Data Collection:** Students earn marks on those specific questions.
3.  **Aggregation:** Marks from different questions testing the same CO are combined.
4.  **Threshold Evaluation:** We check if the student met a "Target" proficiency for that CO.
5.  **Class-Level Analysis:** we count what percentage of the class met the target.
6.  **Attainment Leveling:** We assign a final score (0-3) based on that percentage.

---

## 2. Step-by-Step Process Detail

### Step 1: Defining the COs
Before any marks can be processed, the course must have defined COs (e.g., CO1, CO2, CO3). These are the "buckets" into which performance data will be poured.

### Step 2: Assessment Mapping (The "Linkage")
This is the most critical step. When a teacher creates an assessment (like a Quiz or Mid-Term):
-   They define **Questions** (Q1, Q2, Q3...).
-   They assign **Max Marks** to each question.
-   They **Map** each question to one or more COs.
    -   *Example:* Q1 (5 marks) -> CO1; Q2 (10 marks) -> CO1 & CO2.

### Step 3: Raw Marks Entry
Marks are uploaded via Excel. The system stores these as `MarkScores`, linked to a `Student`, an `AssessmentQuestion`, and indirectly to the `COs` mapped to that question.

### Step 4: Calculating Individual Student CO Performance (Tier 1)
For a specific student and a specific CO (e.g., Student A, CO1):
1.  **Identify Questions:** Find all questions in all assessments mapped to CO1.
2.  **Sum Obtained Marks:** Add up all marks Student A earned on those specific questions.
3.  **Sum Max Marks:** Add up the maximum possible marks for those same questions, **but only for those questions where the student has a recorded mark entry** (i.e., unattempted questions are excluded from the denominator).
4.  **Calculate Percentage:** 
    `Student CO1 % = (Obtained Marks / Max Marks) * 100`

### Step 5: The "Target" Filter
Every course has a **Target Percentage** (e.g., 60%). This is the "passing mark" for a specific outcome.
-   If Student A's CO1 % is **>= 60%**, they are considered to have **"Attained"** CO1.
-   If it is **< 60%**, they have not.

### Step 6: Class-Level Aggregation (Tier 2)
Now we look at the whole class (or section) for CO1:
1.  **Count Successes:** Count how many students "Attained" CO1 (met the 60% target).
2.  **Calculate Success Rate:** 
    `Class Success % = (Number of Students Meeting Target / Total Students) * 100`

### Step 7: Assigning the Attainment Level
Finally, we compare the `Class Success %` against the course's **Attainment Levels** (thresholds).
*Standard NBA Thresholds (Example):*
-   **Level 3 (High):** >= 70% of students met the target.
-   **Level 2 (Medium):** >= 60% of students met the target.
-   **Level 1 (Low):** >= 50% of students met the target.
-   **Level 0 (Not Attained):** < 50% of students met the target.

---

## 3. Detailed Example Calculation

**Scenario:**
-   **Course:** Data Structures
-   **Target:** 60%
-   **Thresholds:** L1: 50%, L2: 60%, L3: 70%
-   **CO1 Mapping:** 
    -   Quiz 1, Q1: 10 Marks
    -   Mid-Term, Q3: 20 Marks
    -   *Total Max for CO1:* 30 Marks

**Student Performance (CO1):**
-   **Student A:** Earned 18/30 (60%). **Met Target? YES**
-   **Student B:** Earned 25/30 (83%). **Met Target? YES**
-   **Student C:** Earned 12/30 (40%). **Met Target? NO**

**Class Result:**
-   2 out of 3 students met the target.
-   `Success Rate = (2 / 3) * 100 = 66.67%`

**Final Attainment:**
-   66.67% is >= 60% (L2) but < 70% (L3).
-   **Final CO1 Attainment Level = 2**

---

## 4. Data Integrity Rules

1.  **Unattempted Questions:** If a student does not attempt a specific question (no mark entry exists), both the obtained marks (0) and the max marks for that question are excluded from their individual CO calculation. This ensures students are not penalized for questions that were not part of their specific attempted set (e.g., optional questions).
2.  **Multiple CO Mapping:** If a question is mapped to both CO1 and CO2, the marks earned on that question contribute fully to the calculation of *both* COs.
3.  **Weightage:** In this system, questions are weighted by their `Max Marks`. A 20-mark question has twice the impact on a CO's attainment as a 10-mark question.
4.  **Internal vs External:** Both internal assessments (Quizzes, MSTs) and external assessments (Final Exams) are aggregated together if they are mapped to the same CO.
