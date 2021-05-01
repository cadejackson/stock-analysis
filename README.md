# Stock Market Analysis for 2017 and 2018

## Project Overview

#### The purpose of this project was to create a workbook with VBA so that Steve could quickly analyze data from 12 stocks for the years 2017 and 2018 so that he could advise his parents on potential investments.

## Results

### Stock Analysis

#### In 2017 all stocks performed well with positive rate of return with the exception of TERP.  9 of the 12 stocks performed exceptionally well with a return greater than 20%.  2018 was much different with 10 of the 12 stocks posting a negative return.  The two stocks that posted a positive return in 2018 were ENPH and RUN.  ENPH performed consistently in 2017 and 2018.  

![2017 Stock Analysis](https://github.com/cadejackson/stock-analysis/blob/main/Resources/2017%20Stock%20Analysis.png) ![2018 Stock Analysis](https://github.com/cadejackson/stock-analysis/blob/main/Resources/2018%20Stock%20Analysis.png)

### Code Performance

#### There was a slight difference between the original code and the refactored code in terms of time; however, there was a 9 - 12 perecent reduction in time.  When executing the code on larger datasets a 9 - 12 percent reduction could make a big difference in the total execution time.  

![Original Code 2017](https://github.com/cadejackson/stock-analysis/blob/main/Resources/Original%20Code%202017.png) ![Refactored Code 2017](https://github.com/cadejackson/stock-analysis/blob/main/Resources/Refactored%20Code%202017.png)

![Original Code 2018](https://github.com/cadejackson/stock-analysis/blob/main/Resources/Original%20Code%202018.png) ![Refactored Code 2018](https://github.com/cadejackson/stock-analysis/blob/main/Resources/Refactored%20Code%202018.png)

## Summary

#### In this project the original VBA code was refactored in order to improve performance.  When deciding whether to refactor or not one should consider the future of the project and if it will get more complex with addtional updates.  Refactored and simplified code can make future updates much easier to implement.  One must also weigh the project goals when deciding whether to refactor or not.  If a product needs to meet a hard deadline then it is advisable to forego refactoring in order to meet the deadline by submittign a working product.  The code could then be refactored at a later time if necessary and if time allows.  

When considering refactoring specifically for a VBA script, a lot of the statements from the previous paragraph still apply.  If you will be executing code on large datasets with hundreds of thousands of rows then it would be a good idea to refactor your code for performance imporvement reasons to save time in the future.  For simple VBA scripts that perform routine tasks that do not take long ot execute or take up much memory it is probably not worth your time to refactor the code.  Since VBA is limited to excel the projects in general will most likely be smaller where refactoring may not be necessary.
