Section 1: Data Cleaning and Preparation
Step 1: Import the Dataset
• Open Excel and go to Data → Get & Transform → Get Data → From File → From Workbook.
• Select the dataset and load it into Power Query.
Step 2: Identify Missing & Invalid Values
• In Power Query, filter columns to find missing or invalid entries.
• Identify null values in critical fields like Region, Sales, Profit, Quantity, and Discount.
Step 3: Data Transformation
• Replace missing values:
➢ For Region → Use "Unknown".
➢ For Sales & Profit → Use the mean or median of the respective column.
➢ For Quantity → Use the mean or median of the respective column.
➢ For Discount → Use the mean or median of the respective column.
• Remove invalid rows that cannot be reasonably corrected.
Step 4: Add a New Column for Profit Margin
• Click "Add Column" → "Custom Column".
• Use the formula:
[Profit] / [Sales] * 100
• If you get null values, check for missing or zero values in the Sales column (to avoid division by zero).
Section 2: Statistical Analysis
After cleaning the data, perform statistical analysis to derive insights.
(a) Calculate Summary Statistics for Sales & Profit
Use Excel Formulas to compute:
• Mean (Average) → =AVERAGE(range)
• Median → =MEDIAN(range)
• Mode → =MODE.SNGL(range)
• Variance → =VAR.P(range)
• Standard Deviation → =STDEV.P(range)
(b) Identify Top-Performing Regions & Customer Segments
1. Use a PivotTable to analyze total Sales & Profit:
➢ Select data → Insert → PivotTable
➢ Drag Region to Rows
➢ Drag Sales & Profit to Values
➢ Sort in descending order to find the top regions & customer segments.
(c) Detect Outliers Using the IQR Method
1. Compute Q1 (25th percentile) & Q3 (75th percentile) using:
=QUARTILE.INC(range, 1) → Q1
=QUARTILE.INC(range, 3) → Q3
2. Compute Interquartile Range (IQR):
= Q3 - Q1
3. Determine Outlier Boundaries:
Lower Bound = Q1 - 1.5 * IQR
Upper Bound = Q3 + 1.5 * IQR
4. Identify outliers:
Use Conditional Formatting to highlight values below Lower Bound or above Upper Bound.
Section 3: Data Visualization
Step 1: Create Key Charts
1. Monthly Sales Trend (PivotChart)
➢ Create a PivotTable with Order Date (Grouped by Month) and Sales.
➢ Insert a Line Chart to show monthly trends.
2. Regional Profit Contribution (Bar Chart)
➢ Create a PivotTable with Region and Total Profit.
➢ Insert a Bar Chart to visualize contributions.
3. Sales vs. Profit Scatter Plot
➢ Select Sales & Profit columns → Insert → Scatter Chart.
➢ Click on the data points → Right-click → Format Data Series → Change marker colors.
Step 2: Add Interactive Slicers
• Click on PivotTable → Go to PivotTable Analyze → Click Insert Slicer.
• Select Region, Product Category, and Customer Segment for dynamic filtering.
