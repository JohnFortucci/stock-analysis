# STOCK ANALYSIS USING EXCEL VBA

## Project Overview

The purpose of this project was to perform a refactoring excercise on some previous written code that analyses stock prices of a limited number of stocks, and ensure that the code is optimized to perform the analysis on a greater number of stocks if required.

## Code Functionality

The original code prompts the user to select the year for analysis and process the respective years stock history data and calculates the yearly return of the stocks and output this in a results sheet. Formating the results to show the year return as a visaul cell formatting Green if the yearly return was positive and red if the yearly return was negative , if the return value is zero the cell will not be colored.

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
![Clear Worksheet Compare Image](/resources/UI_Comparison.PNG)

Based on the above image we can confirm that the original code and the refactored code have the same functionality and no issues have been introduced by the refactoring activity.

### Compare output results

The image below show side by side comparisons of the results of runs of the original code and the refactored code for 2017 and 2018. The data used for the comparison was identiical for the respective years for both original and refactored runs.

![2017 2018 Side by Side Image](/resources/2017_2018_Side_by_Side.PNG)

Based on the above side by side comparison , it can be seen that the :- 
- Tickers , Total Daily Volumes and Return values are match for each year. 
- Color coding in the Return field match for each year
- Headers and formating are identical 

From this we can confirm that the refactored code generated the same output and calculation of the original code and therefore the refactored code generates the same output as the original code.

### Perfomance

One of the objectives of the refactoring process was to improve the overall performance of the original program.

Runtimes are calculated buy capturing the VBA timer value , after the user has selected the respective year , capturing the VBA times value once the code is complete and subtracting the start timer value from the end timer value , this will give the runtime of the code.

The images below show the runtimes of the original code for 2017 (Left Side Image) and 2018 (Right Side Image)
 
![2017  2018 Side by Side Image](/resources/Original_Challenge_2017.png)           ![2017 2018 Side by Side Image](/resources/Original_Challenge_2018.png)

The images below show the runtimes of the refactored code for 2017 (Left Side Image) and 2018 (Right Side Image)
 
![2017 2018 Side by Side Image](/resources/VBA_Challenge_2017.png)           ![2017 2018 Side by Side Image](/resources/VBA_Challenge_2018.png)

Based on the images above the runtimes are :- 

Original code executed with data for 2017   : 0.8632813 Seconds

Refactored code executed with data for 2017 : 0.1601563 Seconds

This shows an improvement in the runtime for 2017 data of 0.703125 which is an improvement of approximately 81%

Original code executed with data for 2018   : 0.882813 Seconds

Refactored code executed with data for 2018 : 0.1875 Seconds

This shows an improvement in the runtime for 2018 data of 0.695313 which is an improvement of approximately 79%

Based on the data above there is an approximately 80% improvement on runrimes with the refactored code.


## Summary


### Refactoring

It is not always possible to write clear , concise and efficient code first time this may be due to but not limited to :
- Deadline contraints 
- Prototyping due to weak requirements

Code refactoring is an activity aimed at improving the overall quality and design of existing code. While code refactoring seems to be a reasonable activity there are however some disadvantages. The perceived adavantages / disadvantages are described below.


#### Advantages of refactoring
Refactoring is aimed at improving curent code , below we have identified some , but not all advantages :- 
- Imporved performance of the code , this can be achieved by removing or consolidating redunant code or loops.
- Improved commenting and formatting , this can be achieved by correctly formatting indents on loop and conditions etc. Inserting clear and concise comments describing various parts of the code. This can make future debugging or development easier for yourself or other developers.as it is much easier to read and understand we formatted and commented code.

#### Disadvantages of refactoring

There are some disadvantages to refactoring , below we have identified some , but not all disadvantages :- 
- Can be time consuming and may not be visible to users , they may see an improvement on runtime , but it would be invisible for them to see an improvement in code readability.
- There is the possibility of introducing errors to previously tested and accepted code.
- Can be percieved as inappropriate use of time , re-inventing the wheel and is detracting from performing new code developments


### Original Code v Refactored code - Advantages and Disadvantages.

Based on the activity we performed by refactoring the Sock Analysis worksheet we can determine the following advantages / disadvantages.

The advantages are that is a definte improvement in code performace , approximately 80% runtime improvement. The code is more readable and well commented this will be helpful if we ever had to revisit the code due to a reported bug or additional functionality is required.

It could be percieved as disadvantages are that the effort invested to increase the perfomance by approximately 0.7 seconds outweighs the amount of developer time required , there was a necessity to retest the code to ensure that the existing functionality and output was retained after refactoring.

