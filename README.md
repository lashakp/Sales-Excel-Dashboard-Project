Excel Sales Dashboard Project – Full Documentation & Analysis Report 
1. Project Overview 
This project involved building a complete interactive Sales Dashboard in Microsoft Excel using 
real-world messy data. The goal was to simulate an end-to-end analytics workflow, including 
data cleaning, transformation, analysis, visualization, and dashboard design. The dashboard 
provides insights into sales trends, product performance, customer activity, and regional sales 
distribution. 
The final output includes: 
• A fully cleaned and structured dataset 
• PivotTables and PivotCharts for analysis 
• Dynamic slicers for interactivity 
• Executive-style KPI cards 
• A polished, user-friendly dashboard 
2. Raw Dataset Description 
The original dataset contained 30+ rows of sales transaction data with the following columns: 
• Date 
• Customer Name 
• Product 
• Quantity 
• Unit Price 
• Region 
• SalesRep 
• Notes 
• Total Sales 
The data intentionally included inconsistencies to mimic real-world conditions, such as: 
• Mixed date formats 
• Numbers stored as text 
• Missing or inconsistent notes 
• Text-based numeric values 
• Varying delimiter styles (slashes, dashes, dots) 
3. Data Cleaning & Standardization 
3.1 Fixing Date Formats 
Dates were inconsistent among formats like: 
• 23/03/2025 
• 11-Mar-25 
• 2025.03.09 
• 3/11/2025 
• 15-03-2025 
A new column Date was created using: 
=DATEVALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE([@Date], ".", "/"), "-", "/"), " ", "")) 
This converted all date variations into a true Excel date serial number. 
Other clean columns were created with helper columns which were later deleted. 
3.2 Numeric Cleanup 
Fields like Qty, Unit Price, and Total Sales contained text-based numbers. 
Recalculated TotalSales accurately with: 
=[@Qty] * [@UnitPrice] 
3.3 Creating an Excel Table 
The data was converted into an Excel Table for: 
• Auto-expanding formulas 
• Cleaner references 
• PivotTable compatibility 
• Better formatting consistency 
4. Feature Engineering 
4.1 Month Column 
To analyze monthly sales trends: 
=TEXT([@CleanDate], "MMM") 
4.2 KPI Metrics 
Computed: 
• Total Orders 
• Total Revenue 
• Average Order Value (AOV) 
• Sales per Product 
• Top Product (Widget A) 
These were turned into dashboard KPI cards. 
5. PivotTables Created 
Built the following PivotTables: 
5.1 Sales by Region 
• Rows → Region 
• Values → TotalSales 
5.2 Sales by Product 
• Rows → Product 
• Values → TotalSales 
5.3 Monthly Sales Trend 
• Rows → Month 
• Values → TotalSales 
• Sorted by month order 
5.4 Sales by SalesRep 
• Rows → SalesRep 
• Values → TotalSales 
6. PivotCharts Developed 
Each PivotTable was visualized using a PivotChart: 
• Line Chart → Monthly Sales Trend 
• Bar Chart → Sales by Region 
• Stacked Column Chart → Region × Product 
Charts were formatted with clear labels, consistent colors, and minimal clutter. 
7. Dashboard Interactivity with Slicers 
Slicers were added for: 
• Date 
• Product 
• Region 
• Customer 
These slicers were linked to all PivotTables, allowing the user to filter the entire dashboard 
dynamically. 
8. KPI Cards Design 
KPI cards were created for: 
• Total Orders 
• Total Revenue 
• Average Order Value 
• Top Product 
• Top Product Sales Value 
Card formatting included: 
• Large, bold numbers 
• Smaller descriptive labels 
• Soft background colors 
• Thick borders 
• Highlighting (e.g., a star next to the top-performing product) 
9. Final Dashboard Layout 
The dashboard was structured into three clean sections: 
Top Section – KPI Summary 
A clean row of KPI cards aligned and evenly spaced. 
Middle Section – Slicers 
Compact, resized slicers arranged horizontally for easy access. 
Bottom Section – Charts 
A combination of line charts, bar charts, and stacked column charts aligned neatly for visual 
clarity. 
10. Skills Demonstrated 
This project demonstrates proficiency in: 
• Data cleaning (dates, numeric fields, structure) 
• Excel formulas (IF, VALUE, DATEVALUE, TEXT, etc.) 
• Excel Tables & structured references 
• PivotTables for aggregation 
• PivotCharts for visualization 
• Dashboard design principles 
• Slicers for interactivity 
• KPI metric creation 
• Layout alignment & visual consistency 
This is a complete demonstration of end-to-end Excel analytics capability. 
11. Final Deliverables 
• Cleaned Excel dataset 
• PivotTable analysis 
• PivotCharts 
• KPI metric calculations 
• Fully interactive Excel dashboard 
• This written project documentation 