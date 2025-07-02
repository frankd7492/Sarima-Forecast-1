This data is 17 months of monthly sales to 2 customers by item. The goal is to forecast the next 6 months of sales for these items and customers. The xlsx file contains the original data and then also the power query/reformatted data.
The CSV is the reformatted data used to load into R

1. Used power Query to reformat the data in Excel.
2. Load the reformatted data into R
3. Filter out discontinued items 
4. Filled in missing values from months with no shipments to be 0
5. Identify outliers on a customer x item x order level, this is calculated by being either  below Q1 -1.5 x IQR or above Q3 + 1.5 x IQR
Where IQR is Q3-Q1, Q1 being the 25th percentile and Q3 being the 75th percentile.
6. Used a SARIMA model to forecast 6 months on a customer x item level which uses regression and detects seasonality on historic data. 
7. Exported the results to Excel and created a pivot table to more closely resemble the historic data in the original file sent.
8. Takeaways: some of the item x customer combinations' forecasts come in pretty flat across the 6 months, this is because of few data 
points since we tend to be shipping on a monthly basis. I notice many months we don't ship, so I would assume that we would aggregate across 
the months we don't ship to the month that we do. Other combinations come in at 0 across the 6 months because of low historical shipments if any at all. 
If any are new products when we would want to assign similar products to be used as a historic reference. In both instances, I put high value on communication 
with the sales and operations teams in order to correctly execute whether aggregating a few months for a larger shipment, or acknowledging the 0 forecast and 
communicating that yes, this is a lower selling item but for one of these months we want to ship 50 (arbitrary).
