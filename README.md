# Stock Analysis 
---
## Overview 

This paper looks at 12 green energy companies stock performances over 2017 and 2018, and ultilises VBA script to determine the total daily volumne and return. It will then provide an analysis on how and why code refracturing is advantageous, as well as providing some insight into challenges and disadvantages of refracturing code both in this example and in general. 


## Purpose 

The aim of this paper is to determine the total daily volumne, and return, for 12 green energy companies, for 2017 and 2018, alongside comparing the execution times of the orignal and the refractored VBA script. This paper aims to determine whether refractoring code allows the VBA script to run faster. 

---

## Analysis 

### Stock Market Performance 
Using both the original code and the refractured code, a comparison of stock performance between the year 2017 and 2018 can be made. Below is a table illustrating the number of trades of the stocks in a given day, which is represented by "Total Daily Volumne', as well as return on investment, 'Return'. 


<p align="center">
<img src="Resources/2017_Stocks.png" width="450">
<img src="Resources/2018_Stocks.png" width="450">
</p>

With the use of conditional formatting in VBA, cell colours turn red when the outcome is a negative return, a green for a positive one, which allows for a simplisitic illustration of the data. It is shown that in 2017, only 1 ticker, ‘Terp’, saw a negative return, in comparrion to 2018, where all but 2 tickers saw a negative turn. 

This data implies that two stocks, ENHP and RUN, would’ve been a good investment as they’re the only two stocks that saw positive returns in both 2017 and 2018. 

### Refractoring Analysis

The second aim of this paper was to refractor the code to loop through all the data once, collecting the same information, and determining whether the refractured code made the script run faster. 

The difference between the original and refractured script was the use nested loops in the original script. A nested loop is a loop inside of a loop, which essentially tells the computer to repeat the code for as many loops that are outlined. The loop in the original script in show below: 

PIC OF NESTED LOOP. 

Here the computer is running through each record of the data set once for each possible ticker. 

By using the code above, 

In order to make this script run quicker, we can refracture the code and remove the nested loop and create and array (see image below). Instead of the computer running through each record of the data set once for each possible ticker, it’ll run through the data set only once overall. 

PIC OF NEW CODE

With this modification, the script now runs in 0.281 seconds for the 2017 Analysis and 0.266 seconds for the 2018 Analysis. 

PIC OF NEW TIMES

## Results 

This analysis demonstrates that well refractured code reduces the time it takes to run the script. This example saw a decrease in 0.172 seconds and 0.258 seconds for the 2017 and 2018 Analysis, retrospectively. 

## Summary 


### Advantages of refractoring code 
When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read

### Disadvantages of refractoring code 
