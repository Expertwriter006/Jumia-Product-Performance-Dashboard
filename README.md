# Jumia-Product-Performance-Dashboard
An Excel-based analysis of pricing, discounts, and customer reviews across 112 Jumia product listings, built to help sellers understand what drives product performance and where pricing or marketing strategy could improve.
## Project Objective
The primary objective of this project was to create an interactive Excel dashboard for analysing the performance of products listed on Jumia.
The final dashboard provides an overview of product performance using key performance indicators (KPIs), charts, product rankings, and category breakdowns.
The dashboard is intended to help Jumia sellers and decision-makers understand how pricing, discounts, ratings and customer engagement interact.
## The Dataset
The dataset contains information and data about Jumia products.
The original dataset was made of **115 product records** and **6 main columns**
1. Product       :          Name of the product
2. Current Price :          Current selling price (In Ksh)
3. Old Price     :          Original price before discount
4. Discount      :          Percentage discount offered
5. Review        :          Number of customer reviews
6. Rating        :          Average customer rating out of 5
7. ## Data Cleaning
The original dataset had several quality issues that I had to address before analysis such as:
- Duplicate records.
- Prices were stored as text containing the `KSh` currency symbol
- Ratings were stored in text such as `4.5 out of 5`
- Review values appeared as negative numbers.
- Missing values in the Review and Rating fields
### Action Taken
I started by checking for **duplicates** in the Original dataset and deleted them.
I did this by selecting the whole data set, and under the **Data** tab in Excel, I clicked on **Remove Duplicates**.
This step reduced the number of products to **111**

Secondly, I changed the values in Original Price field,which were in `Text` form, into `Numbers.`
Similarly, the rating field contained values such as `4.5 out of 5` which is in Text format, and need to be in Numerical values

#### Checking for Missing Values
Missing values can interfere with calculations and visualizations. The original dataset had missing values in the **Review** and **Rating** Columns.

#### Cleaning the Review Column
The Review column had its values written as negative such as `-2, -4, -14, -7.`
Logically, the number of customer reviews cannot be negative. I treated the negative signs as erroneous and removed the negatives.  

#### Creating the Discount Amount Column
One of the calculated fields required by the project was the absolute discount amount.
The formula is:
`=Old Price - Current Price`

#### Creating the Rating Category
I created the Rating Category to group the products according to their customer ratings as follows:

- Poor: Rating below 3
- Average: Rating between 3 and 4
- Excellent: Rating above 4.5
The Excel formula is:
`=IF(F2<3,"Poor",IF(F2<=4.4,"Average","Excellent"))`

#### Creating the Discount Category
I created the discount category in three groups:

- Low Discount: Below 20%
- Medium Discount: 20%–40%
- High Discount: Above 40%
The Excel formula is:
`=IF(D2<20%,"Low Discount",IF(D2<=40%,"Medium Discount","High Discount"))`

### Descriptive Statistics
After the data cleaning and transformation, I calculated the descriptive statistics as follows, with the excel functions:

- Total Products (`=COUNTA(A2:A113)`):**111**
- Avg Current Price (`=AVERAGE(B2:B113)`): **1181.37**
- Avg Old Price(`=AVERAGE(C2:C113)`):**1803.10**
- Avg Discount(`=AVERAGE(D2:D113)`): **37%**
- Avg Rating (`=AVERAGE(F2:F113)`): **3.88**
- Total Reviews(`=SUM(E2:E113)`): **721**
## Correlation Analysis
I investigated three major relationships:
1. Discount and reviews
2. Rating and reviews
3. Price and rating
### 1. Discount Vs Customer Reviews
The first relationship examined was whether products with higher discounts receive more customer reviews.
The correlation between discount percentage and number of reviews, of the final 111 products was **-0.139.**
This indicates a very **weak negative relationship** in the dataset.
The results, therefore, suggests that **higher discounts** are **not associated** with substantially **higher customer engagement** in this dataset.
### Rating Vs Customer Reviews
The correlation for this relationship was **0.066.**
This indicates a **weak positive relationship**.
Therfore, highly rated products do not necessarily receive significantly more reviews.
### Price Vs Rating
The correlation between price and rating is **0.104.**
This represents a very **weak positive relationship.**
Therefore, more expensive products are not necessarily rated substantially higher than cheaper products.
## Creating KPI cards for the Dashboard
The Key Performance Indicators for my dashboard included:

`- TOTAL PRODUCTS 
- AVERAGE PRICE 
- AVERAGE DISCOUNT 
- AVERAGE RATING 
- TOTAL REVIEWS`
## Pivot Tables
### Rating Category
This pivot table was to provide a breakdown of products into:

- Poor
- Average
- Excellent
## Key Business Findings
The analysis produced several important findings as follows:
### 1.High Discounts Do Not Guarantee High Engagement
The correlation between discount percentage and reviews was approximately **-0.139**, indicating a weak negative relationship.
Therefore, increasing discounts does not automatically result in more customer reviews.

### 2.High Ratings Do Not Guarantee High Demand
The relationship between ratings and review gave a correlation of **0.066,** indicating almost no linear relationship. Some highly rated products have very few reviews.
Sellers should, therefore, consider both Rating and Rating Volume to evaluate performance.

### 3.Price Does Not Strongly Determine Rating
The correlation between price and rating of **0.104** indicates a very weak positive relationship.
Therefore, premium pricing does not automatically result in higher customer ratings.

### 4.Some Products Have Strong Engagement but Poor Satisfaction
A clear example in this category is the 120W Cordless Vacuum Cleaner. with 
`29 Reviews
2.8 Rating`
This product attracts considerable customer engagement but has poor satisfaction. It needs to be analysed what could be the cause for this trend.

### Some Products Are Heavily Discounted Despite Poor Ratings
A good example in this category is the 5-PCS Stainless Steel Cooking Pot Set that has:
`55% Discount, 2.1 Rating, 13 Reviews`
This indicates that high discounting does not necessarily solve customer satisfaction problems.

## Recommendations
### 1. Avoid excessive reliance on discounts
The sellers should not assume that increasing discounts will automatically increase customer engagement.
Promotional startegies should involve improvements of product quality and better customer experienvces

### 2. Investigate products with high reviews, and low ratings
Products with large numbers of reviews but low ratings should receive urgent attention. A clear example is the 120W Cordless Vacuum Cleaner

### 3. Analyse customer feedback
The sellers should examine the content of negative reviews to identify recurring problems that may include:
- Quality
- Durability
- Size
- Functionality

### 4. Use ratings and the number of reviews together
A 5.0 rating based on one review should not be treated in the same way as a 4.7 rating based on dozens of reviews.
A better performance framework should consider:

`Customer Rating + Review Volume + Discount + Price`

### 5. Reconsider high discounts on poorly rated products
When a product has high discount and poor rating, the selllers need to examine the underlying problem before enhancing the discount.

## Limitations
- Missing data in ratings and review counts of some products
- The database does not contain a dedicated product-category field. Adding categories would make it possible to compare performance across product groups
- The sales data contains reviews but does not have actual sales quantities
- The revenue and profit-margin data are not available

## Conclusion
This project demonstrated how Microsoft Excel can be used to transform raw e-commerce data into a practical business intelligence dashboard.
The final cleaned dataset contains 111 products after removing three duplicate records and the problematic sofa-cover record.
The analysis demonstrates the importance of cleaning and validating data before creating a dashboard.
The final findings show that:



