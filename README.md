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


#### Refactoring Runtimes
##### Original Code 
![green_stocks2017](https://user-images.githubusercontent.com/95381303/149685157-6888df0b-4e32-43ff-9cdc-17c4e08a2262.png) ![green_stocks2018](https://user-images.githubusercontent.com/95381303/149685158-96ce0a73-b335-4dcd-aa80-3f95a56d5656.png)

##### Refactored Code 
![VBA_Challenge_2017](https://user-images.githubusercontent.com/95381303/149685174-58ec0d94-e0f2-48e5-bfc6-08a4c935942d.png) ![VBA_Challenge_2018](https://user-images.githubusercontent.com/95381303/149685175-326d6148-3c7f-40ce-a3a1-6fd8d9f03f9b.png)

## Summary
#### Refactoring Analysis
 The advantages of refactoring                           
  - reducing the time it takes to run the code
  - clear clutter ; more easily read
  - can expand to include more data making it less specific and more useful         
 
 The disadvantages of refactoring 
  - can possibly create more errors: bug fixes
  - can be time consuming; especially with inexperience










