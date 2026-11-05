# Netflix-Movies-and-TV-Shows
# Netflix Titles – Data Cleaning & Preprocessing in Excel

A data cleaning project on the Netflix Titles dataset (8,807 rows, 12 columns), done entirely in **Microsoft Excel** using built-in tools and formulas.

## Objective

Prepare raw Netflix titles data for analysis by:

- Identifying and handling missing values
- Removing duplicate rows
- Standardizing text values (country names, extra spaces)
- Converting dates to a consistent `dd-mm-yyyy` format
- Keeping column headers clean and uniform (lowercase, no spaces)
- Checking and fixing data types

## Dataset

| Column | Description |
|---|---|
| show_id | Unique ID of the title |
| type | Movie or TV Show |
| title | Title name |
| director | Director(s) |
| cast | Cast members |
| country | Country/countries of production |
| date_added | Date added to Netflix |
| release_year | Year of release |
| rating | Content rating (TV-MA, PG-13, etc.) |
| duration | Length in minutes or number of seasons |
| listed_in | Genres |
| description | Short summary |

## Tools Used

- Microsoft Excel (Filter, Go To Special, Remove Duplicates, Text to Columns, Find & Replace, Power Query)
- Formulas: `TRIM`, `SUBSTITUTE`, `CLEAN`, `IF`, `ISTEXT`, `ISNUMBER`, `IFERROR`, `LEFT`, `MID`, `FIND`, `VALUE`, `DATEVALUE`, `DATE`, `COUNTBLANK`, `COUNTIF`, `SUMPRODUCT`

## Cleaning Steps

1. **Backup** – kept the raw data on a separate sheet and cleaned a copy.
2. **Trim text** – removed leading, trailing and extra spaces with `=TRIM(A2)`. 88 values in `date_added` had a leading space.
3. **Fix misplaced values** – 3 rows had the duration (e.g. "74 min") in the `rating` column; moved them to `duration`.
4. **Handle missing values** – found blanks with Filter and `Go To Special > Blanks`, then filled them with `Ctrl+Enter`.
5. **Remove duplicates** – `Data > Remove Duplicates`. No duplicate rows were found.
6. **Standardize country** – fixed comma spacing with `=TRIM(SUBSTITUTE(F2,",",", "))`.
7. **Convert dates** – converted `date_added` from text (e.g. "September 25, 2021") to a real date using Text to Columns (MDY) and formatted as `dd-mm-yyyy`.
8. **Fix data types** – `release_year` as a whole number, `date_added` as a date. Verified with `ISNUMBER` / `ISTEXT`.
9. **Split duration** – created `duration_value` (number) and `duration_unit` (Minutes / Seasons).
10. **Add helper columns** – `primary_country` using `=IFERROR(LEFT(F2,FIND(",",F2)-1),F2)` and `year_added`.
11. **Final check** – verified row counts and blanks with `COUNTBLANK` / `SUMPRODUCT`.

## Results

| Check | Before | After |
|---|---|---|
| Total rows | 8,807 | 8,807 |
| Missing `director` | 2,634 | 0 (filled with "Unknown") |
| Missing `cast` | 825 | 0 (filled with "Unknown") |
| Missing `country` | 831 | 0 (filled with "Unknown") |
| Missing `rating` | 4 | 0 (filled with "Not Rated") |
| Misplaced duration in `rating` | 3 | 0 |
| `date_added` with extra space | 88 | 0 |
| Missing `date_added` | 10 | 10 (left blank on purpose) |
| Duplicate rows | 0 | 0 |

`date_added` blanks were kept empty on purpose, because inventing a date would make time-based analysis incorrect.

## Files

- `netflix_titles_cleaned.xlsx`
  - `Cleaned_Data` – cleaned dataset
  - `Raw_Data` – original data for comparison
  - `Cleaning_Log` – before/after counts using Excel formulas
  - `Excel_Steps` – step-by-step guide with formulas and menu paths

## Key Learnings

- Always `TRIM` text before converting dates or finding duplicates.
- Fix misplaced values before filling blanks.
- Use **Paste Special > Values** before deleting helper columns, otherwise formulas break (`#REF!`).
- Leave genuinely unknown dates blank instead of guessing.

## Author

Anuradha Ghosh – aspiring Data Analyst
