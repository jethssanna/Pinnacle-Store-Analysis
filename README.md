# Pinnacle-Store-Analysis
Pinnacle Store | E-Commerce Performance Analysis

Tool: Microsoft Excel (Power Query, Pivot Tables, Dashboard Design)
Type: Exploratory Data Analysis and Business Performance Review
Period: January 2022 to December 2023
Prepared for: Head of Commercial, Pinnacle Store  

The Business Problem

Pinnacle Store is a mid-sized international e-commerce retailer selling across six product categories and six global regions. After a strong performance in 2022, the business entered 2023 on a concerning trajectory. Revenue was falling, profit margins were shrinking, product returns were climbing, and the business was consistently missing its regional revenue targets.

Leadership knew something was wrong but did not know where the problem was concentrated, what was driving it, or how severe the financial damage actually was. The Head of Commercial needed answers fast, with clear evidence and actionable direction heading into 2024.

This analysis was commissioned to investigate the decline end to end, identify the root causes, and present findings in a format that leadership could act on immediately.

The Data

The analysis was built across five datasets covering the full 2022 to 2023 period.

The Sales table contained over 8,000 transaction records with order dates, product details, regional information, revenue, cost of goods sold, profit, and discount percentages. The Products table held the full catalogue of 36 products across six categories with cost price, selling price, and margin information. The Customers table covered 1,800 unique customers with acquisition channel, region, age, and join date. The Returns table recorded every product return with reason and refund amount. The Targets table held monthly revenue targets by region for both years, set optimistically at 10% above 2022 actuals for 2023.

The raw data contained several real world quality issues that required resolution before analysis could begin, including duplicate transaction records, revenue values stored as text, inconsistent region name formatting, missing order dates, and return records with no matching order in the sales history. Each issue was identified, documented, and resolved using Power Query.

Data Cleaning and Preparation

All five datasets were loaded into Excel via Power Query and cleaned before any analysis was performed.

Duplicate order records were removed by deduplicating on OrderID rather than all columns, since subtle differences between duplicate rows would have caused a full column deduplication to miss them. Revenue values stored as text were converted to numeric format to ensure accurate aggregation. Region names with inconsistent capitalisation and trailing spaces were standardised using Trim and Capitalize Each Word transformations in Power Query. Rows with missing order dates were removed entirely rather than imputed, since date accuracy is critical for trend analysis and the affected rows represented less than 0.2% of total records. Fifteen return records with no matching OrderID in the Sales table were identified and excluded from category and product level returns analysis, though they were flagged in the documentation as a data quality issue worth investigating at source.

After cleaning, XLOOKUP formulas were used to enrich the Returns table with Category and ProductName from the Sales table, and the Sales table with AcquisitionChannel from the Customers table, enabling cross table analysis without merging the datasets permanently.

Analysis Approach

The analysis was structured around the six key questions in the commercial brief, working from the macro picture down to the specific detail.

The investigation began with the headline numbers — how much did revenue and profit fall and when did the decline start? From there it moved into category and regional breakdowns to identify where the problem was concentrated. Profitability was analysed separately from revenue because a business can lose revenue for market reasons, but margin erosion usually signals an internal decision problem. Returns were investigated both by volume and by financial impact to understand the hidden cost beyond what the revenue figures alone suggested. Acquisition channel performance was analysed to understand whether the customer base was shrinking because of poor retention, poor acquisition, or both.

Pivot tables were used for all analysis, with calculated metrics including profit margin percentage, average order value, return rate, and revenue per customer by channel derived through manual formula columns and XLOOKUP enrichment alongside the pivot table outputs.

Key Findings

The decline was severe and started early. Revenue fell 58.1% from 2022 to 2023, and profit fell 64.2% over the same period. The decline did not build gradually — it began in January 2023 and maintained a consistent downward trajectory throughout the year, suggesting a structural shift rather than a temporary dip.

Every region missed its 2023 target significantly. The Middle East missed by 69.4%, Africa by 67%, Latin America by 65.6%, Europe by 61.4%, North America by 60%, and Asia Pacific by 56.4%. No region came close to its target. Given that targets were set only 10% above 2022 actuals, these misses indicate a business wide collapse rather than isolated regional underperformance.

Electronics drove the most revenue but Books and Media declined the sharpest proportionally. Electronics was the highest revenue category in both years but showed the steepest absolute decline. Books and Media had the lowest revenue overall and the lowest profit margin in 2023 at 39.6%, down from 48.6% in 2022.

Profit margins fell across every category between 2022 and 2023. The overall margin dropped from 54.1% in 2022 to 46.2% in 2023, a decline of 14.7 percentage points. The combo chart analysis showed a clear pattern: categories with higher average discounts in 2023 experienced the steepest margin compression, suggesting the business used discounting to try to stimulate sales but succeeded only in eroding profitability without recovering volume.

The returns problem added a hidden financial burden. 230 returns were recorded in 2023, representing a return rate of 6.4% and generating $21,000 in refund costs. Sports and Outdoors had the highest return count at 47, followed by Home and Living at 42 and Fashion at 41. The most common return reason was Not as Described, which points to a product listing or expectation management problem rather than a quality issue.

Customer numbers fell only marginally but average order value collapsed. Total customers declined just 7.34% from 1,675 in 2022 to 1,552 in 2023, which means Pinnacle Store largely retained its customer base. However average order value dropped 44.2% year on year, from $130 to $86. Customers did not leave — they simply spent significantly less when they did shop. This is a critical distinction because it shifts the recovery strategy away from customer acquisition and toward basket size and spend per visit.

All acquisition channels declined in revenue but Referral remained the most efficient. Referral delivered the highest revenue per customer at $128, followed closely by Email Campaign and Direct. Paid Ads generated the lowest revenue per customer at $120 despite being a paid channel, suggesting poor return on ad spend relative to organic and referral traffic.

Recommendations

Pinnacle Store should prioritise margin recovery over volume recovery in the short term. Continued discounting is compressing margins without driving sufficient volume to compensate. A disciplined reduction in discount rates, starting with the categories showing the highest discount to lowest margin correlation, would stabilise profitability before broader growth initiatives are pursued.

The returns problem in Sports and Outdoors and Home and Living should be investigated at the product listing level. With Not as Described as the leading return reason, the most likely fix is improving product descriptions, images, and size or specification guidance rather than the products themselves.

Given that customer numbers held relatively stable while average order value collapsed, the recovery strategy should focus on increasing spend per customer rather than acquiring new customers. Loyalty incentives, product bundling, and upsell strategies targeted at the existing base are likely to generate faster returns than new acquisition spend.

Referral and Email Campaign channels should receive increased investment given their superior revenue per customer performance. Paid Ads should be reviewed for efficiency, with budget reallocated toward channels demonstrating better return on spend.

The consistent and severe target misses across all six regions in 2023 suggest that the 2024 target setting process needs to be grounded in realistic baseline assumptions rather than optimistic growth projections from a declining base.

Dashboard Overview

The analysis was presented across three interactive Excel dashboard pages.

Executive Summary

The Executive Summary page provides the headline KPIs and the macro story of the decline, including the monthly revenue trend, category and regional performance comparisons, and the revenue versus target table showing the variance by region.

Executive Summary Dashboard

Profitability and Returns

The Profitability and Returns page investigates the internal causes of margin erosion and the hidden cost of returns, including the discount versus margin combo chart, category margin comparison across both years, top product declines, and the full returns breakdown by category, product, and reason.

Profitability and Returns Dashboard

Customer and Channel Analysis

The Customer and Channel Analysis page examines who Pinnacle Store's customers are and how channel performance and regional value have shifted, including revenue and customer count by acquisition channel, revenue per customer by channel, average order value by region, and the monthly customer count trend comparing both years.

Customer and Channel Analysis Dashboard

All three dashboards are interactive, with Year, Region, and Category slicers allowing the viewer to explore the data dynamically across any combination of filters.

Skills Demonstrated

Power Query for data cleaning, transformation, and enrichment across five datasets. Pivot table analysis for multi dimensional breakdowns across categories, regions, channels, products, and time periods. XLOOKUP for cross table data enrichment between datasets. Dashboard design including KPI cards with year on year comparison indicators, interactive slicers, combo charts, and formatted data tables. Business analysis and storytelling through structured findings and recommendations presented in a consulting style report format.

Tools Used

Microsoft Excel, Power Query, Pivot Tables, XLOOKUP

The dataset for this project was purpose built to simulate real world e-commerce business challenges, including intentional data quality issues that required resolution before analysis could begin. All company names, figures, and scenarios are fictional and created solely for analytical demonstration purposes.

