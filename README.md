# VBA-challenge

For this challenge on VBA, I had to first loop through all the worksheets. I created a variable (ws) to reference all worksheets. Inside the for loop, I declared the last row and created another a nested for loop for rows 2 to the last row. 
This way, I was able to check if we are still within the same ticker or not. I was able to find the quarterly change, percent change, and the total stock volume and print them in their respected columns. Through the built in functions on VBA, I was able to color the quarterly change red or green based on the positive or negative change. I closed the nested for loop after this. 
In order to find the greatest % increase, greatest % decrease, and the greatest total volume, I created another nested for loop. I created if statements to loop through all the rows to find the maximum and minimum percent changes as well as the greatest total volume and printed the tickers in their respected cells. 
These are the results of my code on the alphabetical testing and the multiple year stock data.

![image for alphabetical_testing](https://github.com/otybaasandorj/VBA-challenge/blob/main/images/Alphabetical_testing.png)

![image for multiple year stock data](https://github.com/otybaasandorj/VBA-challenge/blob/main/images/Multiple_year_stock_data.png)
