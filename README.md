# Stock-Analysis

# Overview of Project

## The Purpose of this Analysis
To help Steven provide his parents with data-informed recommendations for future stock trading and investments, an analysis was performed on a data set containing daily stock trading information spanning 12 different stock options for 2017 and 2018.

# Results

To help Steven determine whether his parent's initial investment in DQ stock was profitable, our initial analysis examined 2018 returns for that ticker index. The result indicated a yearly loss of -62%. This prompted us to look for other investment options that fared better.

Our next step was to calculate the yearly returns for all of the tickers present in our data set across both years. The results showed that DQ actually did fare better than any other stocks present in 2017, yielding a 199% return.
In 2018 only RUN and ENPH yielded positive year-end returns, and both were greater than 80%. Of the two, ENPH fared much better than RUN in 2017. 

Following this discovery, we wrote a code using VBA script that would allow us to automatically generate a report of the total daily stock volume and year-end investment-return for the 12 tickers across the same data constraints for any given year, using a nested for-loop. We then timed how long it took to generate this report.

A desire to improve our code's efficiency prompted us to refactor our code to see if we couldn't improve the run time. Rather than using a nested for-loop, we created 3 additional arrays of output data to create smaller loops that would only need to loop through the corresponding data sheet once.

- The original code for our 2017 sheet ran at 0.56 seconds.
- The refactored code for our 2017 sheet ran at 0.14 seconds, approximately 75% faster.

    <img width="227" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/104729703/174423282-74d6619c-d364-45a4-bfdc-68fb77418829.png">

- The original code for our 2018 sheet ran at 0.59 seconds.
- The refactored code for our 2018 sheet ran at 0.06 seconds, approximately 90% faster.

    <img width="218" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/104729703/174423710-83e181e8-f7a0-49ec-8874-12a8df06bdaa.png">
    
# Summary

## Code Refactoring: Advantages and Disadvantages

As our data findings indicate, refactoring code can be beneficial becuase it can drastically improve the performance time required by a computer program like Excel to run a code program. While we improved the performance of our original VBA script in this analysis by tenths of a second, with a larger dataset this refactoring could become a much more significant time-saver. Additionally, refactored code can be easier to read, which can make it easier for teams to work with. 

Some disadvantages to refactoring code are that it may sometimes require more lines of code to be written, and may require one to completely rethink one's method for obtaining the desired data and calculations. In the case of our VBA script, our refactored code contained more lines of code than our original, though it generated the same report in less time.
