# STOCK ANALYSIS USING EXCEL VBA

## Project Overview

The purpose of this project was to perform a refactoring excercise on some previous written code that analyses stock prices of a limited number of stocks, and ensure that the code is optimized to perform the analysis on a greater number of stocks if required.

## Code Functionality

The original code prompts the user to select the year for analysis and process the respective years stock history data and calculates the yearly return of the stocks and output this in a results sheet. Formating the results to show the year return a a visaul cell formatting Green if the yearly return was positive and red if the yearly return was zero or negative.

## Refactoring Approach

### Objective

The primary purpose of this refactoring activity is to:- 

- Increase the readability and clarity of code.
- Ensure optimal performance.
- Ensure results are consistent with the original tested code.
- Ensure that no new errors are introduced and original; functionality is retained.

### Aproach

- To ensure redability and clarity of code , we will review and edit the code to ensure that variable names are clear and descriptibe , format code to ensure proper idententation.
- Code will be optimised by reducing any reduncies in the code , the performance will be measured by comparing run times of original code and refactored for identical data sets.
- To ensure refactored results are the same as the original code , we will compare the output results of the original code and the refactored code to ensure results match.
- To ensure no new errors are introduced we will test to ensure the code executes as the original code.

## Results

### Overall Functionality 

The overall functionality was to provide a spreedsheet with two active buttons :- 

Button 1 : Clear Worksheet             : When this is selected the cells in the worksheet 'All Stocks Analysis' will be cleared.

Button 2 : Run Analysis For All Stocks : When this is selected the the user will be prompted to select year to generate the analysis for and then the resulting programe will                                              output the Ticker , Total Daily Volume , Return.

The image below shows the functionality of the original code and the refactored code when Clear Worksheet button is selected.
![Clear Worksheet Compare Image](/Resources/UI_Comparison.PNG)

Based on the above image we can confirm that the original code and the refactored code have the same functionality and no issues have been introduced by the refactoring activity.

### Compare output results

The image below show side by side comparisons of the results of runs of the original code and the refactored code for 2017 and 2018. The data used for the comparison was identiical for the respective years for both original and refactored runs.

![2017 2018 Side by Side Image](/Resources/2017_2018_Side_by_Side.PNG)

Based on the above side by side comparison , it can be seen that the :- 
- Tickers , Total Daily Volumes and Return values are match for each year. 
- Color coding in the Return field match for each year
- Headers and formating are identical 

From this we can confirm that the refactored code generated the same output and calculation of the original code and therefore the refactored code generates the same output as the original code.

### Perfomance

One of the objectives of the refactoring process was to improve the overall performance of the original program
 Original Code                                                              Refactored code
![2017 2018 Side by Side Image](/Resources/VBA_Challenge_2017.png)                        ![2017 2018 Side by Side Image](/Resources/VBA_Challenge_2017.png)
