# Data-cleaning-task1
Data Cleaning of Medical appointment Dataset using Pandas Library in Python. 
# Task 1 - Data Cleaning

## Dataset
Medical Appointments - KaggleV2 May 2016

## What I did
- Removed invalid rows (negative age, wrong dates)
- Fixed column name typos (Hipertension → Hypertension)
- Standardized date formats to YYYY-MM-DD
- Encoded Gender (F→Female, M→Male) and No Show (Yes/No → 1/0)
- Exported cleaned data to Excel with formatting

## Libraries Used
- Pandas
- Openpyxl

## Files
- `KaggleV2-May-2016.csv` — Original raw dataset
- `medical_appointments_cleaning.py` — Cleaning script
- `medical_appointments_cleaned.xlsx` — Final cleaned output
