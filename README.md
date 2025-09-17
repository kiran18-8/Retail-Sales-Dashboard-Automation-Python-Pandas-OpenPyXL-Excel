**Retail Sales Dashboard Automation (Excel + Python)**

This project automates the process of analyzing retail sales data (from the Sample Superstore dataset) and generating a multi-sheet Excel Dashboard with automatic formatting, totals, and highlights.
Features
Reads and processes raw sales data from Excel.

**Creates 4 dashboards automatically:**

Product Wise Sale Dashboard – Sales, Profit, Quantity, Avg Sale per Day, Estimated Sale EOM.

Category Wise Sale Dashboard – Category & Sub-Category level totals with grand total.

Customer Wise Sale Dashboard – Customer-level analysis with totals.

State Wise Sales Dashboard – Aggregated by Country, Region, and State.

**Automatically calculates:**

Category totals

Grand total

Estimated month-end sales

Applies conditional styling:
Grand Total rows → Blue background

Total rows → Yellow background

Adjusts column widths dynamically for readability.

Highlights headers with custom colors for better visualization

**Tech Stack**

Python 3.x

Pandas → Data processing & Excel export

OpenPyXL → Excel formatting (headers, colors, column widths)


**Project Structure**

Retail_Sales_Dashboard/

│── Sample - Superstore.xlsx       
│── Retail_Project_Output.xlsx     
│── retail_dashboard.py            
│── README.md 

**Installation**

Clone the repository or copy project files.

Install required dependencies:
pip install pandas openpyxl

**Usage**

Place your input file (Sample - Superstore.xlsx) in the project folder.

Update the outputfile path in the script:
outputfile = r"C:\Path\To\Retail_Project_Output.xlsx"

Run the script:
python retail_dashboard.py

Open the generated Retail_Project_Output.xlsx to explore dashboards.

**Future Improvements**

📈 Interactive Dashboards: Integrate with Power BI or Tableau for dynamic charts and drill-down reports.

📊 Visualizations in Excel: Add charts (line, bar, pie) automatically using OpenPyXL.

⏱️ Automation: Schedule script execution (via Task Scheduler / cron) for daily or weekly updates.

🗂️ Multi-File Support: Allow processing multiple Excel/CSV input files at once.

🌐 Database Integration: Fetch raw sales data directly from SQL databases or cloud storage.

📧 Email Automation: Auto-send dashboards via email to stakeholders after generation.

**Notes**

Works with any dataset following the Sample Superstore schema.

If your dataset has different column names, update the groupby keys inside the script.

Designed for monthly retail performance analysis.

**Author**

👤 G Kiran Kumar

💼 Data Analyst | Excel Automation
