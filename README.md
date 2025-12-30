# Lookup-Dynamic-Analysis-Database-Modeling
The work focuses on solving complex, real-world data problems using dynamic lookups, two-dimensional analysis, rolling calculations, scenario-style comparison tools, and database functions on large transactional datasets.

## Analysis Summary

### 1 Two-Way Lookups (Products & Sales)
- Completed missing product information using lookup functions
- Retrieved:
  - Product IDs
  - Sub-categories
  - Total quantity sold
- Created calculated lookup fields before referencing
- Improved formula readability using named ranges
- Identified lowest total quantity sold value

### 2 Two-Dimensional Lookups with INDEX & MATCH
- Combined `INDEX()` and `MATCH()` for row-and-column lookups
- Created helper columns using `CONCAT()` to align formats
- Built named ranges for:
  - Sales
  - States
  - Quarters
- Extracted accurate sales values from multi-dimensional tables
- Identified Pennsylvania as the highest sales state in the summary table

### 3 Transposed Tables & Ranking
- Applied `INDEX()` + `MATCH()` on transposed datasets
- Looked up profit values across:
  - States
  - Sub-categories
- Calculated Total Profit per State
- Sorted results to identify lowest-performing state
- Identified Texas as the state with the lowest total profit

### 4 Rolling Period Analysis with OFFSET
- Used `OFFSET()` to:
  - Retrieve historical monthly sales
  - Calculate rolling 3-month performance
- Compared current-year vs prior-year performance
- Determined sales increased by ~22% year-over-year

### 5 Dynamic Period Selection
- Combined `OFFSET()`, `INDEX()`, and `MATCH()` to create dynamic period analysis
- Enabled user-controlled offsets via input cells
- Calculated historical sales dynamically based on user-defined month shifts
- Retrieved Consumer segment sales for offset periods

### 6 Building a Comparison Tool
- Constructed a dynamic comparison framework:
  - Starting month
  - Ending month (calculated dynamically)
- Used named ranges and lookup functions for flexibility
- Compared performance across segments
- Identified Consumer as the higher-performing segment for extended offsets

### 7 Visualization & Interactive Analysis
- Extended comparison tool with:
  - Dynamic period selection
  - Segment drop-downs
- Built a column chart linked to dynamic calculations
- Enabled visual comparison of performance across time periods
- Identified incorrect growth assumptions using visual evidence

### 8 Database Functions – AND Conditions
- Used `DSUM()` to aggregate sales with multiple criteria
- Built structured Criteria Tables
- Applied date-based AND logic on transactional data
- Calculated Consumer segment sales for a specific date range

### 9 Custom Criteria & Counting
- Designed custom criteria tables
- Used `DCOUNT()` to count Order Item IDs meeting:
  - Sales thresholds
  - Geographic constraints
- Demonstrated flexible filtering without helper columns

### 10 OR Conditions with Database Functions
- Applied `DAVERAGE()` with OR logic
- Aggregated profit across multiple geographic conditions
- Calculated average profit for Seattle OR Florida transactions

### 11 Wildcards & Pattern Matching
- Used wildcards for advanced text-based filtering:
  - Included "Book*"
  - Excluded "*Bookcase*"
  - Included Product IDs beginning with "TEC-PH"
- Applied `<>` (not equal) logic
- Calculated qualifying Order Item IDs using database functions

## Key Results & Insights
- Successfully replaced legacy lookup logic with modern Excel methods
- Built reusable, dynamic, and user-friendly analysis tools
- Demonstrated strong control over large transactional datasets
- Created flexible models adaptable to user input without formula rewrites

## Skills
- Advanced Excel lookups & references
- Two-dimensional data modeling
- Dynamic period analysis
- Rolling performance calculations
- Database-style aggregation
- Criteria-based analytics
- Interactive Excel tool design
- Data visualization
