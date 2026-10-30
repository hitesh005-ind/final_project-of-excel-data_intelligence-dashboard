# final_project-of-excel-data_intelligence-dashboard

An enterprise-grade transaction analytics and business intelligence solution developed in Microsoft Excel. Designed as the capstone project, this repository features multi-dimensional customer transaction modeling, dynamic lookup systems, text manipulation, list matching, sensitivity analysis, descriptive statistics, and predictive linear regression.

📌 Executive Summary & Objectives

Analyzing a master dataset of 220+ granular retail transaction records across multi-regional sales channels (North, South, East, West) and core product categories (Electronics, Furniture, Office Supplies), this project converts raw transactional logs into an interactive decision-support system.

Key Objectives Implemented:

📅 Date & Time Tracking: Dynamic period modeling using NOW(), TODAY(), EOMONTH(), and DATEDIF().

🔤 Text Abbreviation & List Matching: Automated customer identity code generation and list reconciliation via INDEX/MATCH boolean logic.

🏆 High-Value Customer Ranking: Dynamic array-driven leaderboard without manual sorting (INDEX, MATCH, LARGE).

🎛️ What-If Sensitivity Modeling: Dynamic discount matrix measuring gross volume impact against net margin compression.

📈 Predictive Regression Analytics: Ordinary Least Squares (OLS) Linear Regression ($Y = \text{Profit}$, $X = \text{Sales}$) generated via Analysis Toolpak.

📊 Descriptive Statistics: Complete distribution metrics across order values.

🎯 Executive Dashboard: Single-screen reporting hub featuring dynamic KPI cards, MoM trend charts, category revenue share, and storytelling insights.

🚀 Key Performance Indicators (KPIs)

Metric

Value

Business Description

💰 Total Revenue Generated

₹27,33,760.00

Cumulative top-line sales volume across 220+ transaction records

📈 Total Net Profit

₹8,20,128.00

Net bottom-line operating earnings

📦 Total Transactions

220 Records

Total volume of order line items logged

🛒 Average Order Value (AOV)

₹12,426.18

Mean sales order size across customer purchases

🥇 Top Customer Segment

High-Value Corporate

Top accounts driving ~45% of total revenue

🛠️ Detailed Formula Implementations & Topics

1. Date & Time Functions

Live Timestamp: Header dynamic timestamping using =NOW() and =TODAY().

Period Boundary:

=EOMONTH(DATEVALUE(Date), 0)


Elapsed Transaction Age (Days):

=DATEDIF(DATEVALUE(OrderDate), TODAY(), "D")


2. Text Manipulation & Customer Code Generation

Standardized Identity Codes: Generates unique shorthand identity tags per row:

=UPPER(LEFT(Customer, 3)) & "-" & TEXT(ROW(), "000")


List Reconciliation & Active Buyer Matching: Compares target client lists against historical rosters:

=IF(ISNUMBER(MATCH(Target_Customer, Active_Roster, 0)), "Active Buyer", "Not Found")


3. Dynamic Customer Leaderboard Ranking (INDEX / MATCH / LARGE)

Ranks top spenders dynamically across 220+ records without manual sorting or VBA:

=INDEX('Final Project Dataset'!G2:G221, MATCH(LARGE('Final Project Dataset'!K2:K221, Rank_No), 'Final Project Dataset'!K2:K221, 0))


4. Predictive OLS Linear Regression (Profit vs Sales)

Multiple R (Correlation): 0.9892 (Extremely strong linear correlation)

R-Square ($R^2$): 0.9786 (97.86% of profit variance explained by sales)

5. Advanced Descriptive Statistics Output

Mean (AOV): ₹12,426.18

Median: ₹10,250.00

Standard Deviation: ₹8,942.50

Range: ₹48,800.00 (Minimum: ₹1,200.00, Maximum: ₹50,000.00)

6. Custom Symbol Indicators & Trend Formatting

Month-over-Month (MoM) sales velocity formatted with visual status symbols:

Positive Momentum: ▲ UP

Negative Velocity: ▼ DOWN

