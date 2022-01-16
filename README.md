# stocks-analysis
## Overview of Project
This project is to analyze the performance of 12 stocks during 2017 and 2018. The VBA code contains two ways of obtaining the analysis reports and allows the user to observe the difference in runtime efficiency for both methods
  
### Purpose
The purpose of the project is to analyze runtime efficiency of two methods for obtaining calculations based off data from an excel sheet using VBA scripts and provide insight into the historical performance of stocks from 12 solar companies during 2017 and 2018.

## Results and Analysis
The VBA script allows the user two methods of obtaining the report of data collected from two worksheets each containing around 3000 data points with eight attributes for each data point. 
The first method creates a few variables to store relevant data to perform calculations and compare to adjacent data points using nested for loops and if statements in VBA. On my local machine I was able to obtain a runtime of about 0.6 second using this method. It is important to note runtime can be affected by CPU speed or RAM size. The script took almost twice as long to run if the local CPU resources where strained to around 97%.
The second method was the one required for this VBA challenge and used arrays to store parts of the data to then perform calculations on from referencing data within the array from a given index. This code ran almost ten times faster at around 0.07 seconds and is much more efficient. The images below contain screen shots from the runtime of these two methods on the data sets from 2017 and 2018 respectively.
![VBA_Challenge_2017](Resources/VBA_Challenge_2017.png)
![VBA_Challenge_2018](Resources/VBA_Challenge_2018.png)

## Summary
Generally storing the data into arrays to then reference and perform calculations on is more efficient in terms of runtime.  If the arrays are larger then there could be concerns about resource limitations within the local machine running the script.  Ways to avoid those types of limitations on extremely large data sets are to be aware of data types and how much active memory it takes to store various data types.
- What are the advantages or disadvantages of refactoring code?
Refactoring code takes time to refractor and analyze to create more efficient code. It can take time away from more senior developers to analyze, rewrite and then test code but is generally worthwhile to make more efficient and more scalable code that can be then repurposed for larger datasets. 
- How do these pros and cons apply to refactoring the original VBA script?
The script was magnitudes more efficient in terms of runtime after refactoring and is more viable for larger datasets.  It is important to note more efficiency could be obtained with further analysis of the code, VBA does not natively support multithreading but there are some work arounds to obtain even faster runtimes that would be appropriate for larger datasets.
