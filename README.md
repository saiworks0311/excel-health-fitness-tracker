````text
# 📊 Personal Health & Fitness Tracker — Excel

## 📌 Project Overview

The Personal Health & Fitness Tracker is a beginner-level Microsoft Excel project created as a practical application of fundamental Excel concepts learned during my Data Analytics course.

The goal of this project was to apply basic Excel concepts such as formulas, functions, relative and absolute references, logical functions, sorting, filtering, and Excel tables to a practical real-world use case.

The workbook allows a user to enter basic personal information, calculate simple health-related metrics, record daily health and fitness activities, compare daily values against predefined targets, and generate a weekly summary.

> Note: This project is intended for educational and Excel-learning purposes. Health calculations are simplified estimates and should not be considered medical advice.

---

## 🎯 Project Objectives

The main objectives of this project were to:

- Practice fundamental Excel formulas and functions
- Understand relative and absolute cell references
- Apply mathematical order of operations
- Use logical functions for simple decision-making
- Work with structured tabular data
- Practice sorting and filtering
- Create a weekly data summary
- Connect information between multiple Excel worksheets
- Build a practical mini-project using Excel

---

## 📁 Workbook Structure

The workbook contains the following main sections:

### 1. Profile & Calculator

Used to enter personal information and calculate:

- BMI
- BMI Category
- BMI-based weight range
- Protein target
- Health and fitness target assumptions

### 2. Daily Tracker

Used to record daily:

- Water intake
- Steps
- Exercise duration
- Protein intake
- Sleep
- Calories

It also compares daily values against predefined targets.

### 3. Weekly Summary

Provides basic descriptive analysis of the recorded data using:

- SUM
- AVERAGE
- MIN
- MAX
- COUNT

### 4. Targets & Assumptions

Contains fixed target values used throughout the workbook, allowing formulas to reference the same target instead of manually entering the value in every row.

---

# 🧮 Health Calculations

## 1. BMI Calculation

BMI is calculated using:

BMI = Weight (kg) / Height (m)²

Since height is entered in centimeters, it first needs to be converted into meters.

### Excel Formula

```excel
=B8/((B7/100)^2)
````

Where:

* B8 = Weight in kilograms
* B7 = Height in centimeters
* /100 = converts centimeters to meters
* ^2 = squares the height

This demonstrates:

* Cell references
* Mathematical operators
* Order of operations
* Parentheses
* Exponents

---

## 2. BMI Category

The project uses a nested IF formula to classify the calculated BMI.

### BMI Categories Used

| BMI Range    | Category    |
| ------------ | ----------- |
| Below 18.5   | Underweight |
| 18.5 – 24.9  | Normal      |
| 25 – 29.9    | Overweight  |
| 30 and above | Obesity     |

### Excel Formula

```excel
=IF(B12<18.5,"Underweight",IF(B12<25,"Normal",IF(B12<30,"Overweight","Obesity")))
```

### How the Formula Works

The formula checks the conditions sequentially:

1. If BMI is below 18.5 → Underweight
2. Otherwise, if BMI is below 25 → Normal
3. Otherwise, if BMI is below 30 → Overweight
4. Otherwise → Obesity

This demonstrates the use of nested IF statements.

---

## 3. BMI-Based Weight Range

For educational purposes, the workbook calculates a weight range corresponding to BMI values of 18.5 and 24.9.

The calculation is based on:

Weight = BMI × Height²

### Minimum Weight Formula

```excel
=18.5*((B7/100)^2)
```

### Maximum Weight Formula

```excel
=24.9*((B7/100)^2)
```

This demonstrates:

* Mathematical calculations
* Order of operations
* Parentheses
* Cell references
* Exponents

> This is a simplified BMI-based range and is not intended to represent an individualized medical recommendation.

---

## 4. Protein Target

A simple educational protein estimate is calculated using:

Protein Target = Body Weight × Protein Factor

### Excel Formula

```excel
=B8*$G$11
```

Where:

* B8 = body weight
* $G$11 = fixed protein factor

The protein factor is stored separately in the Targets & Assumptions section.

This allows the factor to be changed in one place without changing the formula itself.

---

# 🔒 Absolute References

One of the main Excel concepts demonstrated in this project is the use of absolute references.

For example:

```excel
=B8*$G$11
```

The $ signs make G11 an absolute reference.

When the formula is copied to another cell, the reference remains:

$G$11

instead of changing to:

G12, G13, G14, etc.

This is useful when a fixed target or assumption needs to be used repeatedly.

---

# 🔄 Relative References

Relative references were also used throughout the workbook.

For example:

```excel
=IF(B2>=H2,"Target Met","Below Target")
```

When this formula is copied from row 2 to row 3, Excel automatically changes:

B2 → B3

H2 → H3

This allows one formula to be copied down an entire dataset.

### Absolute vs Relative References

| Reference | Example     | Behavior            |
| --------- | ----------- | ------------------- |
| Relative  | B2          | Changes when copied |
| Absolute  | $G$11       | Remains fixed       |
| Mixed     | $G11 / G$11 | Partially fixed     |

---

# 📅 Daily Health Tracker

The Daily Tracker contains structured data for seven days.

### Data Fields

| Column         | Description          |
| -------------- | -------------------- |
| Date           | Date of the record   |
| Water (L)      | Daily water intake   |
| Steps          | Number of steps      |
| Exercise (min) | Exercise duration    |
| Protein (g)    | Daily protein intake |
| Sleep (hrs)    | Hours of sleep       |
| Calories       | Daily calorie intake |

Additional columns contain target values and status calculations.

---

# 🎯 Daily Target Tracking

The workbook compares daily values against predefined targets.

The targets include:

* Water target
* Step target
* Sleep target
* Protein target

The target values are stored separately and referenced from the Daily Tracker.

This demonstrates cross-sheet references and absolute references.

---

# 💧 Water Target

The Daily Tracker retrieves the water target from the Profile & Calculator sheet.

### Formula

```excel
='Profile & Calculator'!$G$12
```

This means:

* 'Profile & Calculator' = source worksheet
* ! = separates the sheet name from the cell reference
* $G$12 = fixed target cell

The formula can then be copied down the column so that every day uses the same target.

---

# 💧 Water Status

The workbook determines whether the daily water target has been met.

### Formula

```excel
=IF(B2>=H2,"Target Met","Below Target")
```

The formula checks:

Actual Water >= Water Target

If TRUE:

Target Met

If FALSE:

Below Target

This demonstrates the IF function and comparison operators.

---

# 👟 Step Target

The step target is brought into the Daily Tracker using a cross-sheet absolute reference.

### Formula

```excel
='Profile & Calculator'!$G$13
```

### Step Status

```excel
=IF(C2>=J2,"Target Met","Below Target")
```

The formula compares:

Actual Steps >= Step Target

---

# 😴 Sleep Target

The sleep target is retrieved using:

```excel
='Profile & Calculator'!$G$14
```

The daily sleep status is calculated using:

```excel
=IF(F2>=L2,"Target Met","Below Target")
```

---

# 🥚 Protein Target

The calculated protein target from the Profile & Calculator sheet is used in the Daily Tracker.

### Formula

```excel
='Profile & Calculator'!$B$16
```

The reference is absolute so every row uses the same calculated target.

### Protein Status

```excel
=IF(E2>=N2,"Target Met","Below Target")
```

---

# 🧠 Logical Functions

The project demonstrates the use of:

* IF
* AND
* OR

These functions allow the spreadsheet to make simple decisions based on the data.

---

# AND Function

The AND function checks whether all specified conditions are TRUE.

Example:

```excel
=AND(B2>=H2,C2>=J2)
```

This checks whether:

1. Water target has been met
2. Step target has been met

Both conditions must be TRUE.

---

# Overall Daily Status

The project combines IF and AND to create an overall status.

### Formula

```excel
=IF(AND(B2>=H2,C2>=J2),"All Targets Met","Needs Improvement")
```

The formula checks multiple conditions before returning a result.

This demonstrates how multiple Excel functions can be nested together.

---

# OR Function

The OR function checks whether at least one condition is TRUE.

The project uses it to create an activity check.

### Formula

```excel
=IF(OR(B2>=H2,C2>=J2),"At Least One Target Met","Neither Target Met")
```

This checks whether either:

* Water target is met

OR

* Step target is met

Unlike AND, only one condition needs to be TRUE.

---

# 🔢 Excel Functions Used

## SUM

SUM adds numerical values.

### Total Water

```excel
=SUM(B2:B8)
```

### Total Exercise

```excel
=SUM(D2:D8)
```

These formulas calculate the total amount recorded during the seven-day period.

---

# AutoSum

The workbook also demonstrates AutoSum.

Instead of manually typing:

```excel
=SUM(B2:B8)
```

AutoSum can automatically create a SUM formula based on the surrounding data.

This demonstrates both the Excel interface and the underlying SUM function.

---

# AVERAGE

AVERAGE calculates the arithmetic mean of a range.

### Average Water

```excel
=AVERAGE(B2:B8)
```

### Average Exercise

```excel
=AVERAGE(D2:D8)
```

### Average Steps

```excel
=AVERAGE(C2:C8)
```

### Average Protein

```excel
=AVERAGE(E2:E8)
```

---

# MAX

MAX returns the largest value in a range.

### Highest Steps

```excel
=MAX(C2:C8)
```

### Highest Protein

```excel
=MAX(E2:E8)
```

This identifies the highest recorded values during the week.

---

# MIN

MIN returns the smallest value in a range.

### Lowest Steps

```excel
=MIN(C2:C8)
```

### Lowest Protein

```excel
=MIN(E2:E8)
```

This identifies the lowest recorded values during the week.

---

# COUNT

COUNT counts cells containing numerical values.

### Days Recorded

```excel
=COUNT(C2:C8)
```

Because the Steps column contains numerical values for each recorded day, the formula returns the number of recorded entries.

---

# 📊 Weekly Summary

The Weekly Summary provides a simple descriptive analysis of the seven-day dataset.

The summary includes:

* Total Water
* Average Water
* Total Exercise
* Average Exercise
* Average Steps
* Highest Steps
* Lowest Steps
* Average Protein
* Highest Protein
* Lowest Protein
* Days Recorded

This demonstrates how basic Excel functions can transform raw data into useful summary information.

---

# 🔃 Sorting

The Daily Tracker is formatted as an Excel Table, which provides built-in sorting functionality.

For example, the Steps column can be sorted:

* Smallest to Largest
* Largest to Smallest

The entire row moves together when sorting, keeping the date and corresponding health data associated with each other.

Sorting allows the user to quickly identify days with the highest or lowest values.

---

# 🔍 Filtering

The Excel Table also provides filtering functionality.

Filtering can be used to display only records that meet selected conditions.

For example, the Exercise column can be filtered to display days with more than a specified number of exercise minutes.

Filtering can help answer questions such as:

* Which days had higher exercise durations?
* Which days had a particular water intake?
* Which days met a particular target?

### Sorting vs Filtering

| Feature   | Purpose                                                        |
| --------- | -------------------------------------------------------------- |
| Sorting   | Changes the order of records                                   |
| Filtering | Temporarily hides records that don't meet a selected condition |

---

# 📋 Excel Table

The Daily Tracker was converted into an Excel Table.

Benefits include:

* Automatic filter dropdowns
* Structured formatting
* Easier sorting
* Easier filtering
* Better organization of data
* Easier expansion when new records are added

---

# 🧩 Excel Concepts Demonstrated

| Excel Concept          | Application in Project                         |
| ---------------------- | ---------------------------------------------- |
| Excel UI               | Workbook design and formatting                 |
| Cell References        | Health calculations                            |
| Relative References    | Daily formulas copied across rows              |
| Absolute References    | Fixed targets and assumptions                  |
| Cross-Sheet References | Connecting calculator and tracker              |
| Order of Operations    | BMI calculation                                |
| SUM                    | Weekly totals                                  |
| AutoSum                | Automatic total calculations                   |
| AVERAGE                | Weekly averages                                |
| MIN                    | Lowest recorded values                         |
| MAX                    | Highest recorded values                        |
| COUNT                  | Number of recorded days                        |
| IF                     | Target status and BMI classification           |
| AND                    | Checking multiple conditions                   |
| OR                     | Checking whether at least one condition is met |
| Sorting                | Ordering records                               |
| Filtering              | Exploring selected records                     |
| Excel Tables           | Structuring daily health data                  |

---

# 📈 Example Analysis

Using the sample seven-day dataset, the workbook can answer questions such as:

### How much water was consumed during the week?

```excel
=SUM(B2:B8)
```

### What was the average daily water intake?

```excel
=AVERAGE(B2:B8)
```

### What was the highest step count?

```excel
=MAX(C2:C8)
```

### What was the lowest step count?

```excel
=MIN(C2:C8)
```

### What was the average protein intake?

```excel
=AVERAGE(E2:E8)
```

### How many days were recorded?

```excel
=COUNT(C2:C8)
```

### Did a particular day meet the water target?

```excel
=IF(B2>=H2,"Target Met","Below Target")
```

These examples demonstrate how Excel can be used not only for storing information but also for performing basic data analysis.

---

# 🛠️ Tools Used

* Microsoft Excel
* GitHub

---

# 📂 Repository Structure

```text
excel-health-fitness-tracker/
│
├── Personal_Health_Fitness_Tracker.xlsx
├── README.md
│
└── screenshots/
    ├── profile-calculator.png
    ├── daily-tracker.png
    └── weekly-summary.png
```

---

# 📸 Screenshots

## Profile & Calculator

![Profile & Calculator](screenshots/profile-calculator.png)

## Daily Tracker

![Daily Tracker](screenshots/daily-tracker.png)

## Weekly Summary

![Weekly Summary](screenshots/weekly-summary.png)

---

# 📚 Learning Outcomes

Through this project, I practiced how to:

* Build a structured Excel workbook
* Convert a practical problem into spreadsheet logic
* Write and understand Excel formulas
* Use mathematical operations and order of operations
* Work with relative and absolute references
* Use logical functions
* Connect data across worksheets
* Summarize datasets using basic statistical functions
* Sort and filter structured data
* Build a simple tracking system using Excel
* Present an Excel project as a portfolio project

Most importantly, this project helped me move from learning Excel functions individually to applying them together in a practical project.

---

# 🚀 Future Improvements

This project represents the beginner version of the tracker.

As I continue learning advanced Excel and data analytics, the project can be expanded with:

* Conditional Formatting
* Data Validation dropdowns
* COUNTIF
* SUMIF
* AVERAGEIF
* IFS
* XLOOKUP
* More advanced formulas
* Pivot Tables
* Pivot Charts
* Interactive dashboards
* Automated exercise recommendations
* More detailed trend analysis
* Power BI integration

The current version intentionally focuses on fundamental Excel concepts learned during the initial stage of my Data Analytics learning journey.

---

# ⚠️ Disclaimer

This project is an educational Excel exercise.

The BMI, weight-range, protein, and fitness-related calculations are simplified examples created for learning spreadsheet formulas and data analysis. They are not medical, nutritional, or professional fitness recommendations.

Personal health decisions should be based on guidance from qualified healthcare or fitness professionals.

---

# 👩‍💻 Author

**Sai Divya Shree**

Aspiring Data Analyst | MCA Graduate | Data Analytics Learner

This project was created as part of my journey to develop practical Excel and Data Analytics skills.

```
```
