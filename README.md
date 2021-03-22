# Green Stock Analysis

## Project Overview

We were approached by a friend named Steve to help analyze a green stock to see if it is worth invensting in for his parents and compare it to other green stocks using data from 2017 and 2018. In order to do this, I needed to use the Visual Basic Appication in Excel to find the different
stock's total daily volume and annual return.  From the data that was analyzed, I was able to make a recommendation as to which stock to invest in for Steve's parents.

### Purpose

The purpose of this project was to create an efficient way in excel through VBA to analyze multiple stocks data.  After running the analysis of these 12 stocks, we also looked at how refactoring can make the analysis more efficient and faster.

## Results

### Refactoring the Code

In order to make my code more efficient, I needed to switch the nesting order of my for loops. To do this,
I created a 4 different arrays; tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices. 
The tickers array was used to establish the ticker symbol of a stock. I matched the other three arrays 
with the tickers array by using a variable called the tickerIndex. 

## Summary on Refactoring 

### General thoughts on Refactoring

The major advantage of refactoring code is that you can make the code much more efficient. The major disadvantage of refactoring code is that you can take existing code that already works (even if less efficient), and potentially create code that is not usable and doesn't work. A best practice to avoid this happening would be to save the original code and run the code as you are refactoring it to test it.

### Refactoring in VBA Script

The major advantage of refactoring code in VBA script is that you can use as much as of the original code as you want to and can put your new code side by side with your old code using different modules. The major disadvantage of refactoring code in VBA script is that if you do not have a strong understanding of the syntax, you will struggle to refactor your code as the syntax matters so much more when trying to make your code more efficient. 

