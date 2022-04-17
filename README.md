# VBA Challenge
## Overview of the project
 The purpose of this project is to refactor code to create a more user-friendly process to analyze stock data. In the original code, nested loops were used to analyze our data set in the year 2017 and 2018.  The code had to be refactored in order to create a process for analyzing any year, not just the two-year window indicated previously, in an efficient manner.
 ## Results
 The summary of results is categorized by the stock’s Ticker, Total daily volume and return percentage (was there an overall gain/loss).  Below are screenshots of the results for years 2017 and 2018, respectively.  For the year 2017, only one stock (TERP) resulted in a loss on investment whereas in 2018, majority of the stocks had a negative return on investment.  The stocks in 2018 that did have a positive ROI, drove results that were over 80% profit.  Overall, these summaries show that stocks performed better in the year 2017 than they did in 2018.
 
 ![Allstocks2017](https://user-images.githubusercontent.com/103154070/163725242-5b2a8fe3-174f-4696-a98c-02bdf66dd464.png)
![Allstocks2018](https://user-images.githubusercontent.com/103154070/163725243-444ccc74-6098-4107-a9fb-1c283a642f71.png)


 
 ### Execution Times 
 To reiterate, the purpose of this project was to refactor code to create a more efficient process when analyzing the data set.  In my original code, I used nested loops.  This process would scan through the dataset analyzing each record once for each possible ticker.  As a result, the processing times for each year would be a little over 1 second.
 
 With my refactored code, I created a list and used the array method.  By doing this, we were able to reduce processing time since the computer was now only accessing each data row once, without needing to rescan each record for every possible ticker.  The results are shown below.
 
 ![VBA_Challenge_2017](https://user-images.githubusercontent.com/103154070/163725197-91a675f9-9017-494c-a8dd-20f8e99ee094.png). ![VBA_Challenge_2018](https://user-images.githubusercontent.com/103154070/163725203-e0fdff0c-0eda-437b-a1d9-16ac6bc22175.png)


 ## Summary
 By refactoring my original code, I was able to create a more efficient process with analyzing this data set.  The overall processing time decreased drastically for each year and the script had a more clean and organized appearance.
 ### What are the Advantages or Disadvantages of Refactoring Code?
 Refactoring code has great benefits when it comes to speed and the reduction of complexity, for it is easier to understand.  Doing this, the code will be fresher, but it does have some disadvantages as well.  Sometimes refactoring code can become time consuming and if an error is made, it could be more challenging to find where that error is.  
 ### How do these pros and cons apply to refactoring the orginal VBA Script 
 Refactoring the original VBA script here decreased the amount of time needed to run the macro, therefore resulting in the code being more efficient.  Also, the appearance of the refocused code was very organized and easy to read.
 
 ![array_VBA](https://user-images.githubusercontent.com/103154070/163725267-6e46378c-3e32-417d-887f-557acf43c70b.png)
