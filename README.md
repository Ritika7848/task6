# task6
📊 Excel Formulas & Functions Fundamentals — Electronics Retail

A portfolio project demonstrating practical Excel formulas on a synthetic electronics retail dataset.

🎯 Objective

Practice and demonstrate the Excel functions commonly used by data analysts:

VLOOKUP
XLOOKUP
IF
SUMIFS
COUNTIFS
TRIM
PROPER
LEFT
RIGHT
📁 Project Files
Solution_1_Electronics_Retail_Excel_Formulas.xlsx — complete workbook
Task_6_Report_Solution_1_Electronics.pdf — project report
Task_6_Report_Solution_1_Electronics.docx — editable report
LinkedIn_Solution_1_Dashboard.png — portfolio visual
LinkedIn_Solution_1_Formulas.png — formula showcase
🧾 Dataset

The dataset contains 60 synthetic electronics-retail transactions and a 10-product lookup table.

Key fields include:
Order_ID, Order_Date, Customer, Region, Product_ID, Units, Unit_Price, Sales, Payment_Method, Channel, and Order_Status.

🧮 Formula Examples
VLOOKUP
=VLOOKUP(E2,Lookup_Data!$A$2:$D$11,2,FALSE)

Used to map a Product ID to the product name.

XLOOKUP
=XLOOKUP(E2,Lookup_Data!$A$2:$A$11,Lookup_Data!$D$2:$D$11,"Not Found")

Used to retrieve the product price with an exact-match lookup.

IF
=IF(H2>=20000,"High Value",IF(H2>=10000,"Medium Value","Standard"))

Used to segment transactions by sales value.

SUMIFS
=SUMIFS(Transactions!$H:$H,Transactions!$K:$K,"Completed")

Used to calculate completed sales.

COUNTIFS
=COUNTIFS(Transactions!$K:$K,"Completed",Transactions!$H:$H,">=20000")

Used to count completed high-value orders.

Text Functions
=PROPER(TRIM(C2))
=LEFT(A2,2)
=RIGHT(A2,4)
📈 Analysis Highlights
Completed sales: ₹789,981.00
Completed orders: 43
Average completed order: ₹18,371.65
High-value completed orders: 12
Top region: North
Top category: Monitors
🔍 Data Quality / Edge Cases
Customer names contain extra spaces to demonstrate text cleaning.
Orders have Completed, Pending and Cancelled statuses.
KPI calculations explicitly filter to completed orders.
Lookup formulas use exact matches.
Nested IF logic demonstrates transparent business-rule segmentation.
🛠 Tools

Microsoft Excel / Google Sheets
