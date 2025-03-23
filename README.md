# FIFA WORLD CUP METRIC VISUALISATION & PREMIERE LEAGUE EDA USING MS EXCEL

## Table of Contents
- [Project Overview](#project-overview)
- [Ask](#ask)
- [Prepare](#prepare)
- [Process](#process)
- [Analyze](#analyze)
- [Results & Insights](#results--insights)
- [Conclusion](#conclusion)

## Project Overview
-This project demonstrates my MS Excel skills by analyzing two datasets. The project follows the Ask, Prepare, Process, Analyze methodology to ensure a structured approach to data analysis.
-For the Fifa World Cup analysis, I did focus my process around using power query/power pivot to clean, analyse and visualise the dataset
-The Premiere League data analysis focuses on showcasing my skills using Excel Formulas to do complex calculations and analysis.

## Ask

### Questions asked:

### FIFA World Cup Dataset
- Who are the highest scoring players?
- Which matchups attracted the largest live audience?
- Which team won the most games?

### Premier League Dataset
- Which teams were the most successful between 1993-2014?
- Which team scored the most goals for each season?
- Which team won the most games against the odds (odds being based on bookmaker odds)?
- Which team would have generated the most profit if we bet on them winning for each game played?

## Prepare

### Dataset Information:
- **Source:** Kaggle
- **Format:** CSV

### Data Cleaning Considerations:
- Filtering for relevant job titles
- Removing records with missing salary data
- Standardizing job location formats

## Process

## World Cup Data Analysis in Excel

### Data Preparation and Cleaning
- **Loaded Datasets**: Imported datasets using Power Query.
- **Cleaned Data**: Replaced missing values (e.g., "-") and converted columns to appropriate data types, including `Datetime`.
- **Sorted Data**: Arranged rows chronologically from oldest to newest based on match date and time.
- **Removed Duplicates**: Identified and eliminated duplicate rows to ensure data integrity.

### Data Modeling and Transformation
- **Created Data Model**: Added tables to Power Pivot for relational analysis.
- **Added Calculated Columns**:
  - Created a new column for **Total Goals** per match.
  - Added formulas to calculate **total goals scored per player per match**.
    ```excel 
=IF(LEFT(TRIM([@[Event.1]]),1)="G",1,0)
+ IF(LEFT(TRIM([@[Event.2]]),1)="G",1,0)
+ IF(LEFT(TRIM([@[Event.3]]),1)="G",1,0)
  ```
- **Defined Measures**: Implemented key measures such as:
  - **Median and Average Goals per Match**
  - **Median and Average Attendance Values**
- **Established Relationships**: Connected match and player datasets using `Match ID` as the key.

### Data Analysis and Visualization
- **Player Performance Analysis**:
  - Split the `Events` column in the Player dataset based on delimiters to analyze goals scored per player.
  - Created a **Pivot Chart** to visualize players with the most goals, with slicers to filter by year and competition stage.

- **Attendance Analysis**:
  - Built a **Pivot Chart** to identify teams with the highest average attendance per year and group stage.

- **Match Outcome Analysis**:
  - Added a conditional column in Power Query to determine the **winner of each match**.
  - Created a **Pivot Table and Chart** to highlight the team with the most wins.
    
- **Visualisation**:
  - Added slicers for years and group stages and connected them to the pivot charts
  - Added Visualisations that dynamically update based on values selected in the slicers



