# Refactoring VBA Code for Stock Analysis 

## Overview of Project
#### Purpose
An excel file was created with macros to assist Steve in his analysis of 12 stocks. Later Steve
discloses that he may want to expand the data in his analysis exponentially. I found that the
vba code initially written works well for those 12 but can be improved upon to accomodate his
data expansion needs.

#### Background
I have been tasked with helping out Steve, a friend and recent finance graduate. Steve is
assisting his parents with analysis on green energy stock for invesment purposes. Although their
interests lies in only 1 particular green energy stock he has selected an additional 11 for
comparison analysis. After some review Steve loves the file but discloses he may want to look
at a larger range of stocks in the future.

## Results
#### Stock Analysis
![Stock_Comparison](https://user-images.githubusercontent.com/95381303/149684607-c833599a-0117-4c01-afb9-96ab1c677882.png)
- Steve's parents wanted to invest in DAQO stock (Ticker:DQ ; higlighted in blue).
	- 'DQ' looked great in 2017 but an investment then would have resulted in a complete loss in 2018.
- 'RUN' with only a small return would have been the best investment as it was the only stock to
  increase exponentially in 2018 on it's return.
- Outside of 'RUN' & 'ENPH', which did tank by over 47% in 2018, all of these options Steve chose
 would not have been good investments based on the range of data.
- Steve was right to believe he would need to potentially expand his dataset.


#### Refactoring Runtimes & Code
##### Original Code 
![green_stocks2017](https://user-images.githubusercontent.com/95381303/149685157-6888df0b-4e32-43ff-9cdc-17c4e08a2262.png) ![green_stocks2018](https://user-images.githubusercontent.com/95381303/149685158-96ce0a73-b335-4dcd-aa80-3f95a56d5656.png)
![image](https://user-images.githubusercontent.com/95381303/149690043-2d96d9a6-8d48-4646-966f-20aed6c4b56f.png)

##### Refactored Code 
![VBA_Challenge_2017](https://user-images.githubusercontent.com/95381303/149685174-58ec0d94-e0f2-48e5-bfc6-08a4c935942d.png) ![VBA_Challenge_2018](https://user-images.githubusercontent.com/95381303/149685175-326d6148-3c7f-40ce-a3a1-6fd8d9f03f9b.png)
![Refactored_Stocks_Code](https://user-images.githubusercontent.com/95381303/149690417-44486163-e6c9-4180-a9c3-561606815be6.png)

## Summary
#### Refactoring Analysis
Refactoring code has its advantages and disadvantages. It can reduce the time it takes to run the code and clear clutter so it can be more easily read. Not only by the coder but others as well. Refactoring can also be used to include a larger dataset, making it less specific and more useful. And at times can be used to fix errors in the original code, optimizing its overall efficiency. The disadvantages can be the opposite where more errors are created and something gets broken in the code. It can also be more time consuming, especially with inexperience. 

In this scenario refactoring was an enhancement. It can be seen by the 'Original and Refactored' runtimes for 2017 and 2018 above, they were reduced by a good amount. The original code had a nested loop which required the script to go through each row of data to identify each ticker over and over. In the new code a ticker index was created with 3 output arrays and the nested loop was removed allowing the code to run faster as the script now only has to go through each row once, cutting down the execution time.         
 









