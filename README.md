# Stock Analysis

In this analysis, 12 different green stocks are analyzed for years 2017 and 2018. Besides measuring there performance YoY and trading volumes, this analysis includes a series of scripts to format, clear, and toggle the data by year.

While seeing the stock data is interesting, the main challenge was to create a script that would ask the user which year he or she would want the analysis performed on. The script (or "macro" in Excel parlance) that was ultimately produced could work for any year, assuming the additional years are formatted the way 2017 and 2018 are. 

The pretext for this analysis is imagining a recent college grad with a degree in finance who is wanting to put his skills to use. This person's first clients could be his parents, for whom he does an analysis of recent stock data. As a result he can create some macros that easily analyze the data, but that script might not work as well if the data set was much larger. Consequently, our finance pro needs to write a script that might be harder to read but performs more quickly and uses less hard-coded values and more variables thereby making the script more flexible. 

The excel file & VBA scripts can be found here - [Green Stock Analysis](https://github.com/carlosjennings1991/stock-analysis/blob/main/VBA_Challenge.xlsm)

---

## Stock Performance

Here are those two years shown side by side

![stocks from both years](https://github.com/carlosjennings1991/stock-analysis/blob/main/Resources/Stocks_2017_and_2018.png)

First glance tells us a few things. 

- **2017 was a pretty good year for these twelve**
- **2018 was a pretty *bad* year for these twelve**
- **Enphase Energry Inc. (ENPH) had the strongest performance**

---

### Stock Performance - The Table Itself

This consists of three columns (the ticker, the total daily volume, and the YoY return). These three values are then shown for all 12 green stocks in question, giving us a table of 36 cells. Let's go over each item in detail. 

#### 1: The Ticker

This is just the abreviation of the company name, and is used to save space. 

![ticker example](https://github.com/carlosjennings1991/stock-analysis/blob/main/Resources/ticker%20example.png)

#### 2: The Total Daily Volume

This is a sum of the number of bought & sold shares for the 252 trading days on the market. It is only (*very*) indirectly related to the actual stock price. The volume is merely an indicator of market activity. As you can see in the first screenshot, the stocks with the highest trading volumes weren't necessarily the best performing. 

![daily return example](https://github.com/carlosjennings1991/stock-analysis/blob/main/Resources/total%20daily%20volume.png)

#### 3: The YoY Return (just "Return" in the table)

This is the difference in the closing price for the stock on the last day of trading for the year in relation to the closing price of the stock on the first day of trading for the year ((last price / first price) - 1)). 

![YoY Return example](https://github.com/carlosjennings1991/stock-analysis/blob/main/Resources/YoY%20return.png)

---

## Original Script Performance

The initial script, found here - [original_script](https://github.com/carlosjennings1991/stock-analysis/blob/main/Resources/Initial%20Script.bas), should be noted for two things. 

- It gets the job done
- It's easy to read

Given readability is very important in coding, this can't be understated. Conceptually the code does this. 

- Ask the user which year to run the analysis
- Create an array of all the ticker indexes
- Create an outer loop (one for each of the 12 tickers)
- Create an inner loop, which updates the total volume, starting price and ending price for each ticker. 
- Exit the inner loop
- Print the three values for each ticker
- Exit the outer loop
- Format the cells (red for negative return, green for positive return)

As for performance, the script does it pretty quickly. 

### 2017 speed: 

![2017 speed](https://github.com/carlosjennings1991/stock-analysis/blob/main/Resources/VBA_Challenge_Original_Code_2017.png)

### 2018 speed
![2018 speed](https://github.com/carlosjennings1991/stock-analysis/blob/main/Resources/VBA_Challenge_Original_Code_2018.png)

It processes 3013 rows of information totalling 24,096 cells in under a second. Not bad. However, it could be better. 

---
## Updated Script Performance

The updated script, found here - [modified script](https://github.com/carlosjennings1991/stock-analysis/blob/main/Resources/Modified%20Script.bas) also gets the job done, and much faster. The below screenshots show the times for the years in question. 

### 2017 speed for updated script:

![2017 speed](https://github.com/carlosjennings1991/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

### 2018 speed for updated script:

![2018 speed](https://github.com/carlosjennings1991/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

The updated script processes everything much faster, roughly 5x as fast as the original. However, the updated code is longer and *more difficult to read*, which can frustrate junior developers and makes transferring the code to someone else more difficult.

---

## Conclusion

The two scripts accomplish their goal, to showcase and format the 12 stocks for two years, and they both do it quickly. However, each has their pros and cons. 
Furthermore, editing (refactoring) code yields mixed results. One one hand, you can improve the code, make it perform better, more elegant, easier to scale. On the other hand, editing code gets messy quickly. A tangled knot of comments and half-finished code can emerge. Fortunately, even the refactored code isn't that confusing to read.

### Conclusion - Original Script

- Pro: It gets the job done
- Pro: It's easy to read
- Con: It takes longer
- Con: It would probably perform poorly with a much larger data set

### Conclusion - Modified Script

- Pro: It gets the job done
- Pro: It's very quick
- Pro: It would probably perform well with a much larger data set
- Con: It's more difficult to read







