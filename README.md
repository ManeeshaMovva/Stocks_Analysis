# Stocks_Analysis

##Overview

Steve loves the workbook I prepared for him. At the click of a button, he can analyze an entire dataset. He wants to take this a step further by expanding the dataset to include the entire stock market over the last few years. Although the code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute. 
Refactoring is a key part of the coding process. When refactoring code, I am not adding new functionality; I just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. In this challenge, I will refactor the code from the original notebook to loop through all the data only once to collect the same information. Then, you’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, you’ll present a written analysis that explains your findings.

## Results
	*insert the two excel tables* 

The stock market changes constantly for various reasons, ranging from supply and demand standards at the time, to housing costs to the cost of living. As you can see above, in 2017, all but one stock performed very well and made profits. However, in 2018, all but two companies performed poorly and had negative returns. Although this may seem shocking, it is not. This is a very niche market and with the stock market being so volatile, this is to be expected. Another parameter we can look at is the Total Daily Volume. If you look at those numbers, you can see the number of trades in each year, for each stock. To make sure the conclusions we come up with using the data are correct, we need to look at large amounts of data to have a good enough dataset to work with. We can conclude that the more the stocks are trading (total daily volume), the higher the value of the stock. 

### Execution Times

<img width="363" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/74915619/122831639-2dd9dc80-d2b8-11eb-8669-976b44eb39a0.png">
<img width="306" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/74915619/122831657-3500ea80-d2b8-11eb-8272-66bc59dbc4e4.png">

Our goal of this challenge was to refactor the data so we could make the code more efficient and be able to use large datasets. I refactored the data by creating arrays, looping through the stock data, reading, and storing all of necessary values, and formatting the sheet so Steve and his parents could view the results easily. After refactoring our code, we cut runtime, making it run faster and be able to work through larger datasets.

## Summary 
### Advantages and Disadvantages of Refactoring Code

In general, there are a lot of advantages to refactoring code. A few advantages are that it makes the code more efficient, takes less time to execute and easier to comprehend. When looking at code, reading through several lines of code could become difficult, especially when debugging and depending on the size of the data, it could take longer for the code to execute. When the code is short, sweet and to the point, it’s not only easy to comprehend for the one that’s coding but also for those that could be using the code for their personal use. 

The disadvantages to refactoring are that it can be very time consuming and depending on the dataset, it could be very difficult. When it comes to refactoring, you must first understand what the original code is doing and then plan out how you can refactor to make it more efficient. This can be very time consuming and as there are larger datasets to work with, there’s more factors you have to take into considering when refactoring. 

### Advantages and Disadvantages of the Original and Refactored VBA Script
The advantages and disadvantages I stated above were the ones I felt as though were evident during the refactoring process. By the end of the project, I saw that the code did become more efficient as it took less time to execute and less steps for us to get to the results. However, I had a hard time figuring out how to refactor the code and took quite a bit of time trying to understand how I could refactor it. The planning stage was quite long for me as I’m very new to VBA. Refactoring, however, helped me to understand some of the nuances of VBA and coding within macros. 
