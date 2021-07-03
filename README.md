# Stock Analysis in Excel using VBA
## Overview
The purpose of this project was to create and refactor a code in Excel that would allow our client (Steve) to analyze a variety of stocks from multiple years at the click of a button in order to determine which one would yield the highest return for his parents. Using VBA, we were able to code a module that allowed for just that.

## Results
The refactored code is like the original in multiple ways - the main difference being that we created a variable called "tickerIndex" that allowed us to access the four arrays we created for this analysis: the "tickers" array which indexed each ticker:

![image](https://user-images.githubusercontent.com/86032451/124360861-196fdb00-dbfa-11eb-8c8b-93d0302d1989.png)

And the three ticker output arrays of "tickerVolumes", "tickerStartingPrices", and "tickerEndingPrices":

![image](https://user-images.githubusercontent.com/86032451/124360892-43290200-dbfa-11eb-92fc-a4ecf8b65d9b.png)

We then initialized the "tickerVolumes" to start at "0", then used conditional logic code to determine tickerStartingPrices and tickerEndingPrices for each ticker by checking if the current row is either the first row or last row for each ticker, respectively:

![image](https://user-images.githubusercontent.com/86032451/124361021-c9454880-dbfa-11eb-9c78-930b77a354f1.png)

Finally, we coded the results to display within the table in the "All Stocks Analysis" sheet for each ticker along with their Total Daily Volume and Rate of Return:

![image](https://user-images.githubusercontent.com/86032451/124361068-1fb28700-dbfb-11eb-9562-7207311e3a7b.png)

The result of this analysis is that the rate of return (change of price from beginning to end of year in terms of percentage) for most stocks in 2017 was positive, while in 2018 most had a negative rate of return, except for "ENPH" and "RUN". As those stocks also had a positive rate of return in 2017 as well it is safe to say that those would be safe investments, or at the very least better to invest in than any of the other tickers listed:

![image](https://user-images.githubusercontent.com/86032451/124361231-fba37580-dbfb-11eb-9232-163495931217.png)

![image](https://user-images.githubusercontent.com/86032451/124361245-137af980-dbfc-11eb-9b19-8adec26a0819.png)

## Summary

### Advantages and Disadvantages of Refactoring Code 
The advantage of refactoring code is that it allows for code to run more quickly and efficiently and provides experience for the person refactoring on how to be a better programmer overall. The only disadvantage of refactoring code is that it can take time, which could be an issue if it is being done on a project that has a deadline coming up quickly. 

### Advantages and Disadvantages of the Original and Refactored VBA Script
While the original VBA script was able to provide the desired results, there were really no advantages over the refactored VBA script. The refactored VBA script, on the other hand, not only got the job more quickly (see below screenshots), but also more efficiently in that it combined two different jobs from the original script (i.e., running the analysis and then formatting it so that the rates of return were color-coded).

Run times for the original script:

![image](https://user-images.githubusercontent.com/86032451/124361455-5db0aa80-dbfd-11eb-9a2d-69c4969507ee.png)

![image](https://user-images.githubusercontent.com/86032451/124361471-6a350300-dbfd-11eb-8e23-324b6d6a8cd2.png)

Run times for the new script:

![VBA_Challenge_2017](https://user-images.githubusercontent.com/86032451/124361530-918bd000-dbfd-11eb-829e-1b27fc6a1250.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/86032451/124361536-96508400-dbfd-11eb-89b5-d94ee140a9c3.png)
