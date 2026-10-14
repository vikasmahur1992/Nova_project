NOVARETAIL EXECUTIVE KPI POWERPOINT GENERATOR
==============================================

OVERVIEW
--------
This project generates a five-slide executive KPI PowerPoint presentation using:

1. A NovaRetail Excel dataset
2. An existing PowerPoint presentation as the visual template
3. The Python script: generate_novaretail_kpi_ppt.py

The output presentation follows the same executive review style as the template while using different KPIs from the supplied commercial, customer experience, and operations datasets.


FILES REQUIRED
--------------
Keep the following files in the same folder:

1. generate_novaretail_kpi_ppt.py
   The Python script that reads the data and creates the presentation.

2. NovaRetail_Data.xlsx
   The Excel workbook containing the three required data sheets.

3. NovaRetail_Q2_2026_Executive_Deck.pptx
   The PowerPoint presentation used as the template.

The generated file will be:

4. NovaRetail_Q2_2026_Alternative_KPI_Deck.pptx


PYTHON REQUIREMENT
------------------
Python 3.9 or a later version is recommended.

Check the installed Python version with:

python --version

On systems where Python is called python3, use:

python3 --version


INSTALL REQUIRED PACKAGES
-------------------------
Open Command Prompt, PowerShell, Terminal, or the terminal in Visual Studio Code.

Install the required packages with:

pip install pandas openpyxl python-pptx matplotlib numpy

If the computer uses python3 and pip3, use:

pip3 install pandas openpyxl python-pptx matplotlib numpy

Package purpose:

- pandas: Reads, filters, and calculates the KPI data.
- openpyxl: Allows pandas to read the Excel workbook.
- python-pptx: Opens the template and creates PowerPoint slides.
- matplotlib: Creates the charts used in the presentation.
- numpy: Supports chart positioning and numerical calculations.


EXCEL WORKBOOK STRUCTURE
------------------------
The Excel workbook must contain the following three worksheets. The worksheet names and column names must match exactly.


SHEET 1: Commercial
-------------------
Required columns:

Quarter
Region
Revenue_USD_M
Units_Sold_K
Returns_Rate_pct
Marketing_Spend_USD_M

Example:

Quarter    Region           Revenue_USD_M    Units_Sold_K    Returns_Rate_pct    Marketing_Spend_USD_M
2026-Q2    North America    465              1120            4.3                 34
2026-Q2    Europe           345              880             4.1                 25
2026-Q2    APAC             330              860             3.6                 24
2026-Q2    Latin America    140              390             8.2                 13


SHEET 2: Customer_Experience
----------------------------
Required columns:

Quarter
Region
NPS
CSAT_pct
Avg_Delivery_Days
Support_Tickets_K

Example:

Quarter    Region           NPS    CSAT_pct    Avg_Delivery_Days    Support_Tickets_K
2026-Q2    North America    55     89          2.9                  18
2026-Q2    Europe           53     88          3.2                  14
2026-Q2    APAC             62     92          2.2                  11
2026-Q2    Latin America    29     74          7.5                  31


SHEET 3: Operations
-------------------
Required columns:

Quarter
Region
Stockout_Rate_pct
Return_Processing_Days
Warehouse_Capacity_Util_pct
Fraud_Incidents

Example:

Quarter    Region           Stockout_Rate_pct    Return_Processing_Days    Warehouse_Capacity_Util_pct    Fraud_Incidents
2026-Q2    North America    3.8                  4                         78                             13
2026-Q2    Europe           4.0                  5                         76                             11
2026-Q2    APAC             8.5                  4                         92                             10
2026-Q2    Latin America    12.5                 10                        74                             30


IMPORTANT DATA RULES
--------------------
1. Do not change the worksheet names.
2. Do not change the required column names.
3. Quarter values must use the format YYYY-QN, for example 2026-Q2.
4. Region values should use these names:

   North America
   Europe
   APAC
   Latin America

5. Each quarter should contain one row for each region.
6. At least two quarters are required because the script calculates quarter-over-quarter changes.
7. KPI columns should contain numeric values rather than text.
8. Percentage fields should contain regular percentage values, such as 92 rather than 0.92.
9. Remove completely blank rows from the Excel sheets before running the script.


HOW TO RUN THE SCRIPT
---------------------
Step 1: Put the Python script, Excel workbook, and PowerPoint template in the same folder.

Example folder contents:

NovaRetail_Project/
    generate_novaretail_kpi_ppt.py
    NovaRetail_Data.xlsx
    NovaRetail_Q2_2026_Executive_Deck.pptx
    README.txt

Step 2: Open Command Prompt, PowerShell, or Terminal in that folder.

Step 3: Run the following command:

python generate_novaretail_kpi_ppt.py --data NovaRetail_Data.xlsx --template NovaRetail_Q2_2026_Executive_Deck.pptx --output NovaRetail_Q2_2026_Alternative_KPI_Deck.pptx

If the computer uses python3, run:

python3 generate_novaretail_kpi_ppt.py --data NovaRetail_Data.xlsx --template NovaRetail_Q2_2026_Executive_Deck.pptx --output NovaRetail_Q2_2026_Alternative_KPI_Deck.pptx

After successful execution, this message will appear:

Presentation created: NovaRetail_Q2_2026_Alternative_KPI_Deck.pptx

The generated PowerPoint will be saved in the current folder.


POWERPOINT OUTPUT
-----------------
The script creates the following five-slide presentation:

Slide 1: Executive KPI Summary
- Latest-quarter units sold
- Revenue generated per marketing dollar
- Total fraud incidents
- Automatically generated executive summary

Slide 2: Demand and Marketing Productivity
- Revenue per marketing dollar by region
- Regional units-sold insights
- Quarter-over-quarter APAC and Latin America comparisons

Slide 3: Customer Service Performance
- CSAT by region
- Support tickets by region
- Delivery-time observations
- Latin America support-ticket share

Slide 4: Operational Risk
- Return-processing days by region
- Fraud incidents by region
- Warehouse-capacity observations
- Regional operational-risk commentary

Slide 5: Recommended Actions
- APAC warehouse-capacity action
- Latin America customer-service recovery
- Fraud-control improvements
- Marketing-spend rebalancing


AUTOMATIC CALCULATIONS
----------------------
The script automatically identifies the latest and previous quarters and calculates:

- Total latest-quarter units sold
- Quarter-over-quarter unit growth
- Total latest-quarter revenue
- Total latest-quarter marketing spend
- Revenue per marketing dollar
- Total fraud incidents
- Revenue per marketing dollar by region
- Latin America support-ticket share
- Regional quarter-over-quarter volume changes

The script uses the latest quarter available in the Commercial worksheet. The same quarter should also be available in the Customer_Experience and Operations worksheets.


HOW THE TEMPLATE IS USED
------------------------
The script opens the supplied PowerPoint template to retain its presentation dimensions and theme information.

The existing template slides are removed, and five new slides are created using the defined NovaRetail colors, typography, positioning, charts, KPI cards, and executive-summary structure.

For the best result:

- Use a 16:9 widescreen template.
- Keep a blank slide layout in the template.
- Keep the template file in PPTX format.
- Do not use a password-protected template.


CHANGING THE OUTPUT FILE NAME
-----------------------------
Change the value after --output.

Example:

python generate_novaretail_kpi_ppt.py --data NovaRetail_Data.xlsx --template NovaRetail_Q2_2026_Executive_Deck.pptx --output My_Executive_KPI_Deck.pptx


USING A DIFFERENT EXCEL FILE
----------------------------
Change the value after --data.

Example:

python generate_novaretail_kpi_ppt.py --data Quarterly_Retail_Data.xlsx --template NovaRetail_Q2_2026_Executive_Deck.pptx --output Quarterly_Retail_Deck.pptx

The alternative workbook must still contain the required worksheet names and columns.


USING A DIFFERENT TEMPLATE
--------------------------
Change the value after --template.

Example:

python generate_novaretail_kpi_ppt.py --data NovaRetail_Data.xlsx --template Company_Executive_Template.pptx --output Company_KPI_Deck.pptx

The alternative template should contain a blank slide layout and should preferably use a 16:9 widescreen format.


COMMON ERRORS AND SOLUTIONS
---------------------------
ERROR: ModuleNotFoundError
Cause:
A required Python package is not installed.

Solution:

pip install pandas openpyxl python-pptx matplotlib numpy


ERROR: FileNotFoundError
Cause:
The script cannot find the Excel workbook or PowerPoint template.

Solution:
- Confirm the filename is correct.
- Confirm the file extension is included.
- Run the command from the folder containing the files.
- Alternatively, supply the complete file path.

Example:

python generate_novaretail_kpi_ppt.py --data "C:\Reports\NovaRetail_Data.xlsx" --template "C:\Reports\NovaRetail_Q2_2026_Executive_Deck.pptx" --output "C:\Reports\NovaRetail_Output.pptx"


ERROR: Worksheet named 'Commercial' not found
Cause:
A required Excel worksheet is missing or has a different name.

Solution:
Rename the worksheets exactly as follows:

Commercial
Customer_Experience
Operations


ERROR: Sheet is missing columns
Cause:
A required column was renamed, deleted, or contains extra spaces.

Solution:
Compare the Excel headers with the required column lists in this README file. Copy the required names exactly.


ERROR: At least two quarters are required
Cause:
The Commercial worksheet contains only one quarter.

Solution:
Add data for at least one earlier quarter because the presentation calculates quarter-over-quarter movement.


ERROR: KeyError for a region
Cause:
A required region is missing or spelled differently.

Solution:
Use the following region names exactly:

North America
Europe
APAC
Latin America


ERROR: ValueError related to numbers or charts
Cause:
One or more KPI columns contain text, symbols, commas stored as text, or blank values.

Solution:
- Convert KPI columns to numeric Excel values.
- Remove currency symbols from cells.
- Remove percentage symbols if the column contains values such as 92 or 4.3.
- Fill or remove missing KPI values.


ERROR: PowerPoint file is open or cannot be overwritten
Cause:
The output presentation is already open in PowerPoint.

Solution:
Close the output file and run the script again, or use a different output filename.


CUSTOMIZATION GUIDE
-------------------
The main design colors are defined near the beginning of the Python script:

NAVY
DARK_NAVY
TEAL
GREEN
RED
ORANGE
BLUE
INK
GRAY
LIGHT_BG

Replace the hexadecimal color values to apply a different company color palette.

The presentation text and recommended actions can be changed inside the slide sections of the generate_deck function.

The region order is controlled by:

REGION_ORDER = ["North America", "Europe", "APAC", "Latin America"]

Update this list if a future dataset uses different regions. If the region list is changed, ensure the Excel workbook contains matching region names.


NOTES
-----
- Charts are generated as temporary PNG images and inserted into the PowerPoint.
- Temporary chart files are automatically removed after the presentation is saved.
- The original template is not overwritten.
- The output presentation is a new PPTX file.
- The script uses an existing blank slide layout from the template.
- Keep a backup of the source Excel workbook and PowerPoint template.
# Nova_project
