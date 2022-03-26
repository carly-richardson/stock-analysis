# VBA of Wall Street

## Performing an analysis on green energy stock data

### To determine which stocks performed the best based on their yearly trading volume and stock market returns

## Results

### Stock Performance
The annual stock returns, on the alternative energy stocks I analyzed, were much higher in 2017 compared to 2018. RUN was the only stock that had a higher return in 2018 verse 2017. RUN also traded at a higher volume in 2018 than it did in 2017. DQ had a 199.4% return on investment in 2017, but was down 62.6% year-over-year at the end of 2018.

AllStocksAnalysis_2017.png

AllStocksAnalysis_2018.png

### Code Performance
The first script I wrote to perform the stock analysis, was able to execute the output faster than the refactored code. The original AllStocksAnalysis() script ran in 0.71875 seconds for 2017 and 0.75 seconds for 2018. The refactored script had an execution time of 0.8203125 seconds for 2017 and 0.8125 seconds for 2018.

![image](https://user-images.githubusercontent.com/100643519/160240680-1320cca7-6a93-4a96-9566-9df6ae7b99fc.png)
VBA_Challenge_2017

VBA_Challenge_2018

## Summary

- The advantage of refactoring code is that you already have the structure and ideas of someone else's code to start from. You are only trying to improve the code and make it run more efficiently. The disadvantage is that you are having to build onto someone else's work, and they may not have approached it in the same way that you would have.

- I do see the benefit of collaborating with others to write a code. In general, I think having combined knowledge, experience, and ideas is going to produce better results. However, because I am new to using VBA, I found it more difficult to understand the script when I had not written it from the beginning.
