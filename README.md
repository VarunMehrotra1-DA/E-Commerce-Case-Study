# [E-Commerce Case Study]

## 📌 Project Overview & Problem Statement
### The Challenge
The attached dataset contains ecommerce orders of a company from multiple countries in 2023. It is a mobile accessories ecommerce company with a global footprint. The company is headquartered in Bangalore in India. The supply chain of the company is very strong: customers are located in 14 countries and they receive orders within a week. 
There are various ways in which a customer can order. An order can be placed using the company’s app, its website, using Whatsapp and also via other sources (for example: dialing in the 24*7 support helpline number). Customers are divided into 5 categories: A, B, C, D, E. 
The sales team in the company is structured into 5 teams: Alpha, Beta, Gamma, Delta and Epsilon. Every team has multiple Sales Managers and a few Sales POCs, typically 3-4. Every sales POC is given individual targets and the sum of the targets of the sales POCs is the target of the Sales Manager.

### Data Source
*Datasets
There are 3 tables within the Excel file: Orders, Customers and Sales Targets.
a.	Orders: it contains data corresponding to orders in 14 countries. The Order ID is unique for every order. One customer can place more than one order. The Order Datetime specifies the timestamp (GMT) when the order was placed. The Order Source specifies the source of the order: App, Website, Whatsapp or Other. Every order has a Sales POC associated with it. The Order Value is in rupees.
b.	Customers: The Customer ID is a unique ID for each customer. The age, gender, country and category which the customer is classified into is given in this table. 
c.	Sales Targets: this table contains the Sales POC → Sales Manager mapping as well as the Sales Team of every Sales POC (Point of Contact) and hence Sales Manager. It also contains the targets that each Sales POC has been given. 

Note that Sales POCs are not limited to any customer or country. This means that a Sales POC can be responsible for orders placed by Customer A in Country X as well as for orders placed by Customer B in Country Y.

### Problem Statement
As an analyst, your goal is to help the company derive insights from the data. These insights will help the company take decisions w.r.t. Orders. Eventually it will help the company increase the number of orders, value of orders and hence revenue. Your manager asks you to perform the following operations and get answers to key questions. 

1.	Import all the tables in Power BI Desktop.
2.	At the start of the new financial year, there was a bug in the system which led to inadequate capturing of the Sales Channel. 
a.	Identify the first and last timestamps when we weren't capturing this data.
b.	Handle these missing values appropriately
c.	Why did you use the above approach for imputing the missing values?
3.	Combine the first name and last name to create a Sales Manager Column in Power Query
4.	Create a Data Model with the three tables
5.	We want to understand the Sales Targets and their completion rates. 
a.	Define three buckets: Target not met, Target Met, Exceeded Target for Sales POCs and plot the three categories on an appropriate chart to see the counts.
b.	Which Sales Teams were given how much total sales targets in total. Plot on a chart to see the comparison. Which sales manager was given the highest target? Did he/she meet that target?
c.	On average, how much short were actual sales from targets for the Sales POCs who did not complete them
d.	Plot the % target reached per Sales Manager in each team
6.	In each country, display which Sales Managers have been most successful 
a.	in terms of sales value 
b.	in terms of number of orders and customers
7.	We want to understand the order patterns to forecast demand better. Create an Orders trend (number of orders as well as value) across months. This will allow us to gauge which months were best for the company from a sales standpoint.
8.	We want to understand which country fared how well in terms of number of orders and average order value. Plot both values by country on a map.
9.	In each category and gender of customers, what % of customers did not place an order. 
10.	Which order source proved to be most used by customers across various countries? Plot an appropriate chart to demonstrate the same.
11.	The category of customers stores a lot of hidden insights. Which Category of customers: 
a.	Has most Indian males 
b.	Is the youngest on average
c.	Has the highest sales 
d.	Has the most number of Sales POCs in the Delta team
Show all this data using Appropriate Charts


---

## 🛠️ Solution Approach

### 1. Data Cleaning & Transformation (Power Query)
1.	We deleted the unwanted blank rows in the dataset sheet of orders.
2.	Then we saw how many missing values we have in order source column by applying filter to that.
3.	Order Datetime was sorted in ascending order and found first and last time stamp and that is 12:49 AM – 11:57 PM
4.	We then replaced all null values to website because null values was less than 1% so we use mode to replace null values.
5.	Then from transform tab we applied merge column feature for combining names of the sales manager, this also removes the previous separated columns from the dataset.


### 2. Data Analysis Expressions (DAX)
Highlight the key metrics you created. Use code blocks to make your DAX readable:
```dax
Eye Icon = " 👁 "
```
```dax
if[Gender] = "M" then "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcTkPCUtgW-Rb_artTzuWD0zKFiNPcCR-Ld-k4sxl9dd-Q&s=10"
else if[Gender] = "F"
then "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQhLp1xkO62guS08MPCZ4-glreKOSr0Zwss32BDQwHIEA&s=10"
else "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQhLp1xkO62guS08MPCZ4-glreKOSr0Zwss32BDQwHIEA&s=10")
```
```dax
%_Customers_Didn't_Order = 
VAR TotalCustomers = 
                    COUNTROWS(Customers)

                    VAR CustomersOrdered = 
                    CALCULATE(DISTINCTCOUNT(Orders[Customer ID]),
                    CROSSFILTER(Orders[Customer ID], Customers[Customer ID],Both))

                    VAR CustomerDidNotOrder = TotalCustomers - CustomersOrdered

                    RETURN
                    DIVIDE(CustomerDidNotOrder, TotalCustomers,0)
```
```dax
Average Shortfall = averagex(filter('Sales Targets','Sales Targets'[Target_Bucket]="Target not met"),'Sales Targets'[2023 Sales Target]-'Sales Targets'[Sales_By_POC])
```
```dax
Sales_By_POC = CALCULATE(SUM(Orders[Order Value]),FILTER(Orders, 'Sales Targets'[Sales POC]=Orders[Sales POC]))
```
```dax
Sales_Difference = 'Sales Targets'[Sales_By_POC] - 'Sales Targets'[2023 Sales Target]
```
```dax
Target Achievement % = DIVIDE(SUM(Orders[Order Value]),SUM('Sales Targets'[2023 Sales Target]))
```
```dax
Target_Bucket = SWITCH(True(),
'Sales Targets'[Sales_By_POC]<'Sales Targets'[2023 Sales Target], "Target Not Met",
'Sales Targets'[Sales_By_POC]='Sales Targets'[2023 Sales Target], "Target Met",
'Sales Targets'[Sales_By_POC]>'Sales Targets'[2023 Sales Target], "Exceeded Target"
)
```
```dax
Sales_Manager_Target_Completion = 
SUMMARIZE('Sales Targets','Sales Targets'[Sales Manager], 
"Total Sales", SUM('Sales Targets'[Sales_By_POC]), 
"Total Target", SUM('Sales Targets'[2023 Sales Target]), 
"% Target Completion", SUM('Sales Targets'[Sales_By_POC])/SUM('Sales Targets'[2023 Sales Target]))

---

## 📊 Final Dashboard Visualization
* 1.	It was found already connected in power BI as Sales Target to Orders Table as one to many and Customers to Orders Table as one many because one sales manger can have multiple orders and on customer can place multiple orders.
2.	We created sales by POC column by using calculate function and summing up the orders value and matching sales POC for unique values in order table from sales target table by using filter function inside the calculate. (DAX)
3.	Then we created a bucket column for target not met, target met, exceeded target by using switch statement as there were multiple conditions to state true by comparing Sales By POC of to target sales value.
4.	Plotted a pie chart to visualize the target bucket with count of values which concludes that there are not any targets that met, targets were not met with a count of 37(40.22%) or the targets were exceeded with a count of 55(59.78%).
5.	Plotted a column chart to visualize the total sales of each sales team and this concludes that Alpha team has the highest sales of 3.7M.
6.	Sales Manager Target Completion table was created with columns as Sales Manager, Total Sales (sum of sales by POC), Total Target (sum of sales target), % Target Completion ([total sales/total target]*100)
7.	By Clustered Column chart visual you can see Valarie Nelson (SALES POC) was behind the target by $223,165.
8.	By Combination Chart (Line & Column) we can conclude that Noah Meier had the highest Sale, and was able to achieve the sales 108% approx. of the target and exceeded the target by 8% approx.
9.	As we plotted the matrix for finding the most successful sales manager in each country the results are on the basis of sales values, count of orders and customers we found Benjamin Chen in Austria, Germany, Singapore & Switzerland AND Noah Meier in Australia, Denmark, Finland, France, India, Ireland, Italy, Norway, Spain & USA as the most successful sales manager.
10.	As we plotted line chart to understand the orders pattern, we found that February & November were the best for the company from sales standpoint because these months only have the higher order value amongst the months.
11.	In map Bubble size represents the count of order ID and color intensity represents the average order value, so according to this the fared country found is USA as it has most profitable and high performing markets. Customers here purchase frequently and are comfortable making high-ticket or bulk purchases.
12.	By matrix visual we can conclude that in all categories and considering both genders we have an overall 41.88% of customer’s who didn’t ordered (calculated as total customers - customers ordered divided by total customers). In which category E has the highest percentage due to high pricing, lack of product information or complicated logistics. If we speak of any particular gender so category E has the equal distribution and rest categories shares a nominal difference.
13.	By stacked column chart visual we can conclude that the APP is the order source from where the mostly customer orders across various countries.
14.	In pie chart we can say that customer category A is of 42.31% males, category B is of 39.42% males & category C is of 18.27% males.
15.	By line chart visual we can say that customer category D is youngest on average as 53.03 years.
16.	By column chart visual it is founded that customer category D has done the highest sales value of 3103650. (Doubt on currency/ how to check that)7.	By Bar chart visual we can conclude that Category D has the highest count of Sales POC of 119 from delta team.
*

### Executive View
![Executive Dashboard](screenshots/executive_view.png)

### Regional Performance View
![Regional Dashboard](screenshots/regional_view.png)

---

## 💡 Key Insights & Business Recommendations
* **Insight 1:** There are not any targets that met, targets were not met with a count of 37(40.22%) or the targets were exceeded with a count of 55(59.78%).
* **Insight 2:** Alpha team has the highest sales of 3.7M.
* **Insight 3:** Noah Meier had the highest Sale, and was able to achieve the sales 108% approx. of the target and exceeded the target by 8% approx.
* **Insight 4:** Valarie Nelson (SALES POC) was behind the target by $223,165.
* **Insight 5:** Benjamin Chen in Austria, Germany, Singapore & Switzerland AND Noah Meier in Australia, Denmark, Finland, France, India, Ireland, Italy, Norway, Spain & USA as the most successful sales manager.
* **Insight 6:** February & November were the best for the company from sales standpoint because these months only have the higher order value amongst the months.
* **Insight 7:** It was founded that SPAIN has the highest average of order values even with less count of orders.
* **Insight 8:** We have an overall 41.88% of customer’s who didn’t ordered (calculated as total customers - customers ordered divided by total customers). In which category E has the highest percentage due to high pricing, lack of product information or complicated logistics. If we speak of any particular gender so category E has the equal distribution and rest categories shares a nominal difference.
* **Insight 9:** In Major Countries it was found that APP source was majorly used
* **Insight 10:** Customer category A is of 42.31% males, category B is of 39.42% males & category C is of 18.27% males.
* **Insight 11:** Customer category D is youngest on average as 53.03 years.
* **Insight 12:** Customer Category D has the highest count of Sales POC of 119 from delta team.
* **Insight 13:** Customer category D has done the highest sales value of 31,03,650.

## 💡 Key Performance Indicators
* **Indicator 1:** Overall Actual Sales is 12.28M (sum of order value)
* **Indicator 2:** Total Sales Target were 12.81M  (sum of 2023 sales target)
* **Indicator 3:** Target Achievement % was 95.83% (total sales divided by total target)
* **Indicator 4:** Average Shortfall was 21.07K (sales target - sales by POC from filtered data of target not met)
