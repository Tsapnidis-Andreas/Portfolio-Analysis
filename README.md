# Project Title
## Portfolio Analysis with Python <br>
# Contents
[Info](#Info)<br>
[Portfolio Analysis and Optimization](#Portfolio-Analysis-and-Optimization)<br>
[Constructing The Efficient Frontier and Drawing the CML](#Constructing-the-Efficient-Frontier-and-Drawing-the-CML)<br>
[Disclaimer](#Disclaimer)
# Info
## Programming Language: 
Python <br>
## Libraries used:
[Pandas](https://pandas.pydata.org/#:~:text=pandas%20is%20a%20fast,%20powerful,%20flexible)<br>
[Numpy](https://numpy.org/)<br>
[Matplotlib](https://matplotlib.org/)<br>
[Yfinance](https://pypi.org/project/yfinance/)<br>
[Datetime](https://docs.python.org/3/library/datetime.html)<br>
[Scipy](https://scipy.org/)<br>
[Openpyxl](https://pypi.org/project/openpyxl/#:~:text=openpyxl%20is%20a%20Python%20library%20to)<br>
[Xlsxwriter](https://pypi.org/project/XlsxWriter/#:~:text=XlsxWriter%20is%20a%20Python%20module%20for)<br>
# Portfolio Analysis and Optimization
## How to use
Install the necessary libraries mentioned above<br>
Execute the program <br>
Select portfolio analysis or portfolio optimization <br>
Enter the tickers of the stocks (and the corresponding weights in case you selected portfolio analysis)<br>
Press 'OK'<br>


## How it works
### Portfolio Analysis
The code entails: <br>
Pulling data from yahoo finance via the yfinance Python library <br>
Calculating the daily returns for each stock <br>
Calculating the average daily return and the daily standard deviation of each stock <br> 
Annualizing these values <br>
Calculating the portfolio's expected return and standard deviation<br>
Calculating the sharpe ratio<br>
Obtaining the portfolio's beta and alpha through a linear regression(the S&P500 index serves as the market portfolio)<br>
Saving all the data as an excel file <br>

### Portfolio Optimization 
The code entails:<br>
Pulling data from yahoo finance via the yfinance Python library <br>
Calculating the daily returns <br>
Calculating the average daily return and the daily standard deviation of each stock <br>
Annualizing these values <br>
Identifying the optimal portfolio ie the one that maximizes the sharpe ratio <br>
Analyzing the optimal portfolio as described above <br>


## Example
![Στιγμιότυπο οθόνης 2024-09-04 123142](https://github.com/user-attachments/assets/69da3ea1-f497-4e99-8026-5f9ec7599f8a)
![Στιγμιότυπο οθόνης 2024-09-04 123307](https://github.com/user-attachments/assets/4eb45087-06e5-44fa-8a77-92b3c48b10f8)
![Στιγμιότυπο οθόνης 2024-09-04 123331](https://github.com/user-attachments/assets/608f7380-1aa3-432c-bde6-dc03457c24e7)


# Constructing The Efficient Frontier and Drawing the CML

## How to use
Install the necessary libraries mentioned above<br>
Initialize:<br>
The list containig the tickers of the stock<br>
The risk free rate<br>
The path(the directory where the excel and the png files will be saved)<br>
Execute the program <br>

## How it works
The code entails:<br>
Pulling data from yahoo finance via the yfinance Python library<br>  
Calculating the daily log returns for each stock<br>
Calculating the average daily return and the daily standard deviation for each stock<br>
Annualizing the average return and the standard deviation<br>
Generating random portfolios each with a different combination of weights<br>
Plotting the portfolios<br>
Identifying the portfolio with the lowest standard deviation for each level of return<br>
Plotting the optimal portfolios and constructing the efficient frontier<br>
Identifying the optimal portfolio ie the one with the highest sharpe ratio<br>
Constructing the Capital Market Line<br> 
Saving the optimal portfolios as an excel file<br> 
Saving the chart as a png file<br> 

## Example
This is the output after running the program for the following stocks:<br>
msft<br>
jpm<br>
ko<br>
![1](https://github.com/user-attachments/assets/600424e1-b575-4ff7-b44b-41f84f6825cf)

![2](https://github.com/user-attachments/assets/bb242a76-465c-4181-8099-f1e545a15b3e)




and assuming the risk free rate is equal to 2%<br>

# Disclaimer
This project serves educational purposes only<br>
Under no circumstances should it be used as an investing tool

