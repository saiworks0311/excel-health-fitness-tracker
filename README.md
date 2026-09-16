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
