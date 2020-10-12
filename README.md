# stock-analysis

### Project Overview
The purpose of this project is to get a better understanding of the provided stock market data. Being able to analyze this data will help Steve make a data-driven decision when he helps his parents find a profitable stock. At the moment Steve has no efficient way to process this stock market data. Steve could do this by hand but it would take a large amount of time to process every single row. Steve is fortunate that Microsoft Excel has the programing language Visual Basic for Applications (VBA) which makes repetitive calculations in an extremely short amount of time. In this project, VBA plays a big role in processing about 3000 rows in about .59 seconds. Allowing Steve to focus on analyzing the results and making comparisons within each stock.


### Results

The provided stock data comes from Steve who wants to help his parents better understand the stocks they own and find a higher-yielding stock. Steve is interested in finding the daily volume and yearly return for each stock. Daily volume is the total number of shares traded throughout the day; it measures how actively a stock is traded. The yearly return is the percentage difference in price from the beginning of the year to the end of the year. Steve has the right data but not an efficient way to process the data. This is where VBA reduces the overall time spent making those calculations so Steve can spend his time looking at the results rather than doing repetitive manual computation. Below are the processing times for the year 2017 and the year 2018. As you can see VBA is much faster than any person and possibly any group of people. There are about 3000 rows for each year that contain stock data to process. The stock data contains Ticker, Date, Volume, etc...

#### Processing Time
![2017 Processing time](Resources/VBA_Challenge_2017.png)

![2018 Processing time](Resources/VBA_Challenge_2018.png)

#### Stock Results
From the results below 2017 was a positive year for all stocks except TERP which had a -7.2% return for the year. ENPH, SEDG, and DQ are the three stocks with the highest return for the year 2017. Let's say Steves parents invested in DQ which in 2017 had the highest return overall, DQ had a 199.4% yearly return. 

![2017 Results](Resources/2017_Results.png)

2018 was a negative year for all stocks except ENPH and RUN. ENPH had an 81.9% return and RUN had an 84% return for the year 2018. DQ had a -62.6% year return which is the highest negative year return compared to the other stocks in this data set. Hopefully, Steve can persuade his parents to find a stock that does not have the highest year return but has a positive year return for both years. Based on the data ENPH would be the best stock to go for because both 2017 and 2018 have positive high returns. ENPH does not have the overall highest return compared to other stocks, however, it has a constant positive return for the years 2017 and 2018. 

![2018 Results](Resources/2018_Results.png)

### Summary

#### Advatages/Disadvantages refactoring code
The two main advantages of refactoring code are having dynamic code and efficiency. Before refactoring code, the for loops inside the macro were not as dynamic which then forces values to be reinitialized after every loop iteration. As you can see in the images below the original for loop had to reinitialize the value totalVolume while in the refactored for loop the values for totalVolume were stored in an array. By appending values to an array there would be no need to reinitialize values and the array contains totalVolumes that can be indexed by ticker.

![for loop](Resources/forLoop.png)

![Refactored for loop](Resources/refaForLoop.png)

For this macro refactoring code reduced processing time. This can be seen when comparing the processing time for before and after refactoring. Below the average processing time for the macro before refactoring was about 0.76 seconds. After refactoring the code became much faster because the average processing time was about 0.59 seconds. 

![2017 Processing Time Original](Resources/2017_Original.png)

![2018 Processing Time Original](Resources/2018_Original.png)

That being said two disadvantages of refactoring the code are limiting array lengths and code complexity. For this particular example, there were only 12 different tickers that could be used. Let's say two more tickers are added to the list of tickers. If someone would try to execute the macro, the macro would fail because the arrays are set to only hold 12 values. Below is a snippet of the initialized arrays that can only hold 12 values.

![Arrays](Resources/arrays.png)

Furthermore, by refactoring code the macro increases in complexity. This means that it's difficult to follow along unless there is detailed documentation explaining each line of code. Without detailed documentation, it would almost be impossible to understand what the person who wrote the code was trying to achieve. 


#### Advatages/Disadvantages refactoring VBA script

The main advantages of refactoring the VBA script are code readability and processing time. Documenting each line of code makes it easier for other analysts to understand and make changes to the VBA script. Documenting is as important as the script outputting the correct data. For example, let's say the script works perfectly, however, there is no documentation to follow what the script does. Depending on the complexity of the code it can be almost impossible to understand someone else's code without a guide. Furthermore, refactoring the script improved processing time and this can be seen in the images under processing time. The difference was about 0.20 seconds which is not much. However, let's say Steve wants to look at the year 2019 that has 3000 more rows than 2017 and 2018. The difference between processing time would increase because the factored VBA script is overall much faster and allocates memory better. All things being considered the refactored VBA script is larger than the original VBA script and refactoring scripts is a time-consuming task. The refactored script is about twice as long as the original script. The refactored script has about 135 lines of code while the original script 67 lines of code. Moreover, the refactoring process is not easy because now it becomes a challenge to improve the working script by either making it more dynamic or improving the processing time.
