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

## World Cup Data Analysis in Excel (Power Query/Power Pivot skills focused)

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
<img src="https://drive.google.com/uc?id=1jzxDlVjHaTcAdV2cOeFkJiFpxBzHAdux" alt="World Cup Data Analysis" width="600">

## Premiere League Analysis (Project focus is to showcase use of complex formulas)

### Data Preparation and Cleaning
- **Loaded Datasets**: Downloaded dataset from Kaggle.
- **Cleaned Data**: This dataset did not need to be cleaned.
- **Sorted Data**: Created a new sheet named "Analysis_sheet", where I created a list of unique values for both the teams and the seasons using formulas such as: 
=UNIQUE(PremierLeague!B2:B10000). Also created a second sheet named "Prem_season_analysis" where I included data validation dropdown for the seasons and added the questions I am aiming to answer.

## Identifying the winner for each season
### Breakdown:
1. **Counts home wins (`"H"`)** → Multiplies by **3 points**.  
2. **Counts away wins (`"A"`)** → Multiplies by **3 points**.  
3. **Counts away draws (`"D"`)** → Multiplies by **1 point**.  
4. **Counts home draws (`"D"`)** → Multiplies by **1 point**.  

Each `COUNTIFS` function filters matches for a specific team (`Analysis_sheet!B2`) and season (`season` variable) from the `PremierLeague` sheet. The formula then sums up the total points based on match results.

```excel
=SUM(
    COUNTIFS(PremierLeague!F:F,Analysis_sheet!B2,PremierLeague!J:J,"H",PremierLeague!B:B,season)*3,
    COUNTIFS(PremierLeague!G:G,Analysis_sheet!B2,PremierLeague!J:J,"A",PremierLeague!B:B,season)*3,
    COUNTIFS(PremierLeague!G:G,Analysis_sheet!B2,PremierLeague!J:J,"D",PremierLeague!B:B,season)*1,
    COUNTIFS(PremierLeague!F:F,Analysis_sheet!B2,PremierLeague!J:J,"D",PremierLeague!B:B,season)*1
)
```

- **Show winner on Prem_season sheet**:
```excel
=XLOOKUP(MAX(Analysis_sheet!C2:C52),Analysis_sheet!C:C,Analysis_sheet!B:B)
```
## Identifying relegated teams for each season
### Breakdown:
This formula **concatenates the names of three teams** based on specific conditions:  
1. **The two teams with the lowest positive values in Column C**  
2. **The team with the minimum value in Column M that does an additional check to see if more than one team had the same score, who scored more goals**  

### **Formula**
```excel
=CONCAT(
    XLOOKUP(SMALL(FILTER(Analysis_sheet!C:C, Analysis_sheet!C:C > 0), 1), Analysis_sheet!C:C, Analysis_sheet!B:B) & ", " & 
    XLOOKUP(SMALL(FILTER(Analysis_sheet!C:C, Analysis_sheet!C:C > 0), 2), Analysis_sheet!C:C, Analysis_sheet!B:B) & ", " & 
    XLOOKUP(MIN(K:K), Analysis_sheet!E:E, Analysis_sheet!B:B)
)

=FILTER(Analysis_sheet!$C:$E,Analysis_sheet!$C:$C=SMALL(FILTER(Analysis_sheet!$C:$C, Analysis_sheet!$C:$C > 0), 3))
```
### **Formula to show results**
```excel

=CONCAT(
    XLOOKUP(SMALL(FILTER(Analysis_sheet!C:C, Analysis_sheet!C:C > 0), 1),Analysis_sheet!C:C,Analysis_sheet!B:B) & ", " &
     XLOOKUP(SMALL(FILTER(Analysis_sheet!C:C, Analysis_sheet!C:C > 0), 2),Analysis_sheet!C:C,Analysis_sheet!B:B) & ", " &
     XLOOKUP(MIN(K:K),Analysis_sheet!E:E,Analysis_sheet!B:B) )
```
## Identifying the team with the highest average goals per game
### Breakdown:
This formula **calculates the average of goals scored for the chosen season and corresponding team** and returns **0 if the team did not play in the chosen season**.

### **Formula**
```excel
=IFERROR(
    AVERAGE(
        FILTER(PremierLeague!H:H, (PremierLeague!F:F=Analysis_sheet!B2) * (PremierLeague!B:B=season)), 
        FILTER(PremierLeague!I:I, (PremierLeague!F:F=Analysis_sheet!B2) * (PremierLeague!B:B=season))
    ), 
    0
)
```
### **Formula to show results**
```excel
=CONCAT(
    XLOOKUP(MAX(Analysis_sheet!E2:E52),Analysis_sheet!E:E,Analysis_sheet!B:B) & " with " &
     ROUND(MAX(Analysis_sheet!E2:E52),2) & " average goals per game")
```
## Identifying which team won the most games against bookmaker odds
### Breakdown:
1. **First I created additional columns showing if the Home or Away team had the lower odds (=IF(IF(AC5238>AA5238,G5238,F5238)=F5238,"H","A")), and another column to see if this team won (=IFERROR(IF(IF(IF(AC5238>AA5238,G5238,F5238)=F5238,"H","A")=J5238,"Y","N"),"NAN"))**  
2. **I then created a formula to check how many times the given team won against the odds in the given season selected**

### **Formula to calculate number of games won against the odds**
```excel
=SUM(
  COUNTIFS(PremierLeague!B:B,season,PremierLeague!F:F,Analysis_sheet!B2,PremierLeague!AD:AD,"H",PremierLeague!AE:AE,"Y"),
  COUNTIFS(PremierLeague!B:B,season,PremierLeague!G:G,Analysis_sheet!B2,PremierLeague!AD:AD,"A",PremierLeague!AE:AE,"Y"))
```
### **Formula to show results**
```excel
=IF(
  NUMBERVALUE(LEFT(season,4))<2003,
  "No Odds Yet",
  CONCAT(XLOOKUP(MAX(Analysis_sheet!F2:F52),Analysis_sheet!F:F,Analysis_sheet!B:B)&" won "&MAX(Analysis_sheet!F2:F52)&" games against the odds"))
```

## Identifying which team was the best to bet on
### Breakdown:
1. This formula calculates **how much money we would have made** for the given team in the given season **if we would have bet £10 on them winning every time they played**. Also considering that bookmaker would refund half the bet if the game is a draw.

### **Formula to calculate earnings**
```excel
=SUM(
    SUMIFS(PremierLeague!AA:AA, PremierLeague!B:B, season, PremierLeague!F:F, Analysis_sheet!B2, PremierLeague!J:J, "H") * 10,
    SUMIFS(PremierLeague!AC:AC, PremierLeague!B:B, season, PremierLeague!G:G, Analysis_sheet!B2, PremierLeague!J:J, "A") * 10,
    COUNTIFS(PremierLeague!B:B, season, PremierLeague!F:F, Analysis_sheet!B2, PremierLeague!J:J, "D") * 5,
    COUNTIFS(PremierLeague!B:B, season, PremierLeague!G:G, Analysis_sheet!B2, PremierLeague!J:J, "D") * 5
) 
- COUNTIFS(PremierLeague!B:B, season, PremierLeague!F:F, Analysis_sheet!B2, PremierLeague!J:J, "A") * 10
- COUNTIFS(PremierLeague!B:B, season, PremierLeague!G:G, Analysis_sheet!B2, PremierLeague!J:J, "H") * 10
```
### **Formula to show results**
```excel
=IF(NUMBERVALUE(LEFT(season,4))<2003,"No odds yet",CONCAT(XLOOKUP(MAX(Analysis_sheet!G2:G52),Analysis_sheet!G:G,Analysis_sheet!B:B)))
```
