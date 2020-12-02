# Stock Analysis

In this analysis, 12 different green stocks are analyzed for years 2017 and 2018. Besides measuring there performance YoY and trading volumes, this analysis includes a series of scripts to format, clear, and toggle the data by year.

While seeing the stock data is interesting, the main challenge was to create a script that would ask the user which year he or she would want the analysis performed on. The script (or "macro" in Excel parlance) that was ultimately produced could work for any year, assuming the additional years are formatted the way 2017 and 2018 are. 

The pretext for this analysis is imagining a recent college grad with a degree in finance who is wanting to put his skills to use. This person's first clients could be his parents, for whom he does an analysis of recent stock data. As a result he can create some macros that easily analyze the data, but that script might not work as well if the data set was much larger. Consequently, our finance pro needs to write a script that might be harder to read (for a programmer) but performs more quickly and uses less hard-coded values and more variables thereby making the script more flexible. 

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

This consists of three columns (the ticker, the total daily volume, and the YoY return). These three values are then shown for all 12 green stocks in question, giving us a table of 36 cells. 
