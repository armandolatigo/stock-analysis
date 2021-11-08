# Overview of Project

#### The purpose of this analysis was to help Steve find the total daily volume and yearly return for each stock so he can guide his parents in investment options.  It the trade activity for each stock and the percentage change for each stock from the beginning of the year to the end of the year. 

# Results

#### 2017’s yield in return had a far more favorable outcome than 2018, most of the stocks return in 2018 decreased in rate of return from the previous year, RUN and ENPH are the only two to show growth in the two year timeframe making them the prime candidates for investment. The runtime between the original refactored code is about 0.50 seconds.

![challenge1](https://user-images.githubusercontent.com/92451164/140676428-bb2174f3-1015-47d0-9fc7-e9b4056ac33d.png)
![challenge2](https://user-images.githubusercontent.com/92451164/140676447-ad7b7ca4-f524-4955-9c3c-789f4ac94ce5.png)

![VBAimage1](https://user-images.githubusercontent.com/92451164/140676450-8a1e67eb-4d8e-4f58-91cc-bffb0d0262db.png)


# Summary

#### Well the most obvious difference between the original and refactored code is the runtime difference.  The refactored yields a result faster than the original sequence. The refactored code shows optimization in run time and also using the variable tickerIndex calls from the array reduces the need for creating more variables to pull from the different values each ticker has in the dat set. Using tickerIndex allows to call several data points using the one tickerIndex variable. In all earnest, this seems like ‘cleaner’ code for lack of a better term, or lack of extensive knowledge of VBA at this point.  It runs faster, there’s less code, it eats up less RAM, even thought the two are nearly identical.  The only disadvantage that comes to mind is coming up with the concept of optimizing the code.  I could be, and probably am, wrong here but there is no clear cut disadvantage to refactoring the code.  It’s very similar, runs faster, uses less computing power.  In my (very limited knowledge) opinion.  It’s an optimized version of the previous code. 
