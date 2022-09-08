# Analysis of Green Energy Stocks from 2017-2018

## Overview of Project

### This analysis aims to provide the client, Steve, a fast and efficient Excel macro to analyze a large volume of stocks from any year of interest. This was done so by refactoring code to produce a program that could complete a stock analysis in a shorter amount of time. The output of the program details all stocks, their total daily return, and ROI for a user-inputted year and formats the results to be easily-readable.  


## Results

### Comparison of stock performance between 2017 and 2018
After running the analysis programs on stocks from 2017 and 2018, it's clear that stocks in 2017 performed much better than stocks in 2018. Cells highlighted in green represent a positive ROI, while cells highlighted in red represent a negative ROI. In 2017. only one stock, TERP resulted in a negative return (7.2% loss) whereas 10 stocks in 2018 had negative ROIs, the worst-performing stock DQ resulting in a 62.6% loss. It's hard to believe DQ was the best performing stock n 2017 with almost a 200% return on investment!

#### 2017 Green Energy Stock Analysis
<img width="228" alt="2017 Stock Analysis" src="https://user-images.githubusercontent.com/10901980/189036612-75180639-1b26-4928-8bf8-2d9a457e6abb.png">

#### 2018 Green Energy Stock Analysis
<img width="232" alt="2018 Stock Analysis" src="https://user-images.githubusercontent.com/10901980/189036628-e042abc6-8a20-46c2-9baa-097a8060dc69.png">

If had to choose one stock to recommend investing in, it would be ENPH. It was by far the best performing stock, with positive ROIs in both 2017 and 2018, 130% and 82%, respectively. 

### Comparison of original code vs refactored code
The largest difference between the module 2 code and the refactored code is that the module 2 code utilizes a nested for loop to loop through the stock (year) sheet multiple times to get through analyzing all the stocks.

<img width="338" alt="Original nested loop" src="https://user-images.githubusercontent.com/10901980/189031241-53446e0b-12b6-4c56-bab5-096e8e6c80a6.png">

As you can see in the screenshot, the code is looping through every row in the stock sheet 12 times which means that there are a lot of rows that are unnecessarily looped over before the program can complete the desired analysis. In contrast, the refactored code only loops through the stock sheet one time; it can do this because it keeps track of which stock block it's currently on through a tickerIndex and updates values in a separate arrays accordingly. While an additional for loop is necessary to then access and output the stored values after all rows in the stock sheet has been looped over, it still dramatically cuts down on run-time.

- Arrays to store values for individual stocks as analysis occurs
   
    ```
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    ```
- Additional array to access and output values for each stock
    ```
    For i = 0 To 11

        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1

    Next i
    ```
    
### Comparison of original code vs refactored run-time
As mentioned above, the original 2 code loops through the stock sheet rows 12 times more than the refactored code. This causes the refactored code to have a significantly reduced run time; it reduces the run-time by more than 2/3 of its original run-time.

#### Original run-times for 2017 and 2018:
<img width="251" alt="2017 RT" src="https://user-images.githubusercontent.com/10901980/189033588-b009bae1-6820-479f-bb52-a9bab09f8378.png">
<img width="253" alt="2018 RT" src="https://user-images.githubusercontent.com/10901980/189033601-078480a6-2fc9-464c-ac33-6056af8ec009.png">

    
#### Refactored run-times for 2017 and 2018:
<img width="252" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/10901980/189033700-3f417979-16fc-49e4-99a5-fc7ab22df01e.png">
<img width="252" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/10901980/189033710-bf624245-2c49-4994-9a3d-a3373a0a836f.png">

## Summary

### What are the advantages or disadvantages of refactoring code?
Some advantages of refactoring code are that the core of program is there, so it may save one time in building certain features, and it can improve the quality and run-time of a code. The disadvantage, however, is that it can be difficult and time-consuming as not-the-original author of the code to fully understand what parts of the code do what and edit them appropriately. Trying to maintain the original structure of code could lead to odd work-arounds and introduce new bugs into the program. 

### How do these pros and cons apply to refactoring the original VBA script?
The pro of refactoring the original VBA script was that many features I needed to use were already implemented, such as the array that held the stock names or the formatting function. This saved me some time, and additionally my edits to the code ultimately saved the program user on run-time as well, making the code more efficient. Not too many disadvanges were encountered, other than the fact that I had to take time to read through the entirety of the original code before starting my edits. 
    
