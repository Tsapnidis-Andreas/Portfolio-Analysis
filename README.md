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
Install the necessary libraries stated above<br>
Initialize:<br>
The list containig the tickers of the stock<br>
The risk free rate<br>
The path(the directory where the excel and the png files will be saved)<br>
Execute the program <br>

## How it works
### The code entails:
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
![1](https://github.com/user-attachments/assets/ffde93e5-3168-491f-97c6-e3ea0b115367)
![2](https://github.com/user-attachments/assets/1f9b97e0-031e-4cd7-8590-1c1f3113942f)
![3](https://github.com/user-attachments/assets/0db5e409-d0f3-4e4b-8b36-c6f70927bb96)
![4](https://github.com/user-attachments/assets/0b0d0f2f-706b-4cfc-93ed-34e08feaf5a2)

# Constructing The Efficient Frontier and Drawing the CML

## How to use
Install the necessary libraries stated above<br>
Initialize:<br>
The list containig the tickers of the stock<br>
The risk free rate<br>
The path(the directory where the excel and the png files will be saved)<br>
Execute the program <br>

## How it works
### The code entails:
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

and assuming the risk free rate is equal to 2%<br>

![1](https://github.com/user-attachments/assets/561ccc80-a41f-432a-aa0f-5d3f247c2a7b)
![2](https://github.com/user-attachments/assets/a9c4f19d-8058-441f-93d0-22febcbffbef)
