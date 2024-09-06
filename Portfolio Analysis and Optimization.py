import pandas as pd
import numpy as np
from pandas_datareader import data as wb
import matplotlib.pyplot as plt
import yfinance as yf
import datetime as dt
from datetime import datetime, date
import statsmodels.api as sm
from scipy import stats
import openpyxl
import xlsxwriter
from tkinter import *
from scipy.optimize import minimize


def disappear(a):
    a.place(x=0,y=0,width=0,height=0)

def mode_1():
    for i in objects:
        disappear(i)
    label1.place(x=20, y=80, width=200, height=20)
    label2.place(x=240, y=80, width=200, height=20)
    entry1.place(x=20, y=140, width=200, height=20)
    entry2.place(x=20, y=200, width=200, height=20)
    entry3.place(x=20, y=260, width=200, height=20)
    entry4.place(x=240, y=140, width=200, height=20)
    entry5.place(x=240, y=200, width=200, height=20)
    entry6.place(x=240, y=260, width=200, height=20)
    button1.config(command=OK_1,text='RUN')
    button1.place(x=300,y=300,width=60,height=20)
    global mode
    mode=1
def mode_2():
    for i in objects:
        disappear(i)
    label1.place(x=20, y=80, width=200, height=20)
    entry1.place(x=20, y=140, width=200, height=20)
    entry2.place(x=20, y=200, width=200, height=20)
    entry3.place(x=20, y=260, width=200, height=20)
    button1.config(command=OK_2,text='RUN')
    button1.place(x=300,y=300,width=60,height=20)
    global mode
    mode=2

def OK_1():
    t1=entry1.get()
    t2=entry2.get()
    t3=entry3.get()
    w1=float(entry4.get())
    w2=float(entry5.get())
    w3=float(entry6.get())
    for i in objects:
        disappear(i)
    stocks=[t1,t2,t3]
    weights=[w1,w2,w3]
    analyze(stocks,weights)

def OK_2():
    t1=entry1.get()
    t2=entry2.get()
    t3=entry3.get()
    for i in objects:
        disappear(i)
    stocks=[t1,t2,t3]
    optimize(stocks)

def analyze(stocks,weights):
    weights=np.array(weights)
    global inflation_rate
    stock_data = pd.DataFrame()
    average_daily_returns=[]
    annualised_returns=[]
    average_annual_returns=[]
    for i in stocks:
        stock_data[i],average_daily_return,annualised_return,average_annual_return = obtain_returns(i)
        average_daily_returns.append(average_daily_return)
        annualised_returns.append(annualised_return)
        average_annual_returns.append(average_annual_return)
    average_daily_returns=np.array(average_daily_returns)
    annualised_returns=np.array(annualised_returns)
    average_annual_returns = np.array(average_annual_returns)

    stock_data.columns = stocks
    portfolio_df = calculate_portfolio_returns(weights, stock_data)
    portfolio_df.columns = ['portfolio']
    portfolio_average_daily_return=average_daily_returns.dot(weights)
    portfolio_annualised_return=np.dot(annualised_returns,weights)
    portfolio_annualised_return_per_cent=portfolio_annualised_return*100
    portfolio_average_annual_return_per_cent=np.dot(average_annual_returns,weights)*100
    portfolio_stddev,correlation = calculate_portfolio_risk(weights, stock_data)
    portfolio_df_per_cent=pd.DataFrame()
    portfolio_df_per_cent['portfolio']=portfolio_df['portfolio']*100

    cov_matrix = stock_data.cov()
    sharpe = -calculate_sharpe_ratio(weights,annualised_returns,cov_matrix)
    sp_daily_returns,sp_average_daily_return,sp_annualised_return,sp_average_annual_return =obtain_returns('^GSPC')
    sp_data=pd.DataFrame()
    sp_data_per_cent=pd.DataFrame()
    sp_data['daily returns']=sp_daily_returns
    sp_annualised_return_per_cent=sp_annualised_return*100
    beta,r_square=linear_regression(sp_data['daily returns'].tail(-1),portfolio_df['portfolio'].tail(-1))
    alpha,capm_value=calculate_alpha(portfolio_annualised_return_per_cent,sp_annualised_return_per_cent,beta)
    results = pd.DataFrame()
    results['metric'] = ['average daily return %','annualised return % ','average annual return % (nominal)','average annual return % (inflation adjusted)','standard deviation %', 'sharpe ratio', 'beta', 'R^2','alpha']
    results['value'] = [portfolio_average_daily_return*100,portfolio_annualised_return_per_cent,portfolio_average_annual_return_per_cent,portfolio_average_annual_return_per_cent-inflation_rate, portfolio_stddev, sharpe, beta,r_square,alpha]

    global mode
    print(results)
    if mode==1:
        saving_1(portfolio_df, results,correlation)
    else:
        global optimal_weights
        optimal_weights = pd.DataFrame(optimal_weights)
        optimal_weights.index = [stocks]
        optimal_weights.columns = ['weight %']
        optimal_weights['weight %'] = optimal_weights * 100
        saving_2(portfolio_df, results,correlation,optimal_weights,stocks)
    updating()


def optimize(stocks):
    global optimal_weights
    stock_data = pd.DataFrame()
    avg_returns = []
    for i in stocks:
        stock_data[i],average_daily_return,annualised_return,average_annual_return = obtain_returns(i)
        avg_returns.append(annualised_return)
    avg_returns = np.array(avg_returns)
    cov_matrix = stock_data.cov()

    constraints = {'type': 'eq', 'fun': lambda weights: np.sum(weights) - 1}
    bounds = [(0, 1) for i in range(len(stocks))]

    initial_weights = np.array([1 / len(stocks)] * len(stocks))

    optimized_results = minimize(calculate_sharpe_ratio, initial_weights, args=(avg_returns, cov_matrix), method='SLSQP',
                                 constraints=constraints, bounds=bounds)
    optimal_weights = optimized_results.x
    sharpe=calculate_sharpe_ratio(optimal_weights,avg_returns,cov_matrix)
    analyze(stocks,optimal_weights)


def obtain_returns(stock):
    global n
    global date_index
    end_date = datetime.now().strftime('%Y-%m-%d')
    start_date = (datetime.strptime(end_date, '%Y-%m-%d') - dt.timedelta(days=n*365)).strftime('%Y-%m-%d')
    data = yf.download(stock, start=start_date, end=end_date)
    data=pd.DataFrame(data)
    date_index=data.index
    data['returns']= (data['Adj Close']-data['Adj Close'].shift(1))/data['Adj Close'].shift(1)
    data['returns']=data['returns'].dropna()
    average_daily_return=data['returns'].mean()
    data.index=range(0,len(data))
    a=data.loc[0,'Adj Close']
    b=data.loc[len(data['Adj Close'])-1,'Adj Close']
    total_return=(b-a)/a
    annualised_return=((1+average_daily_return)**250-1)
    average_annual_return=((1+total_return)**(1/n)-1)
    return(data['returns'],average_daily_return,annualised_return,average_annual_return)


def calculate_portfolio_returns(weights,stock_data):
    portfolio_df=pd.DataFrame(stock_data.dot(weights))
    return(portfolio_df)

def calculate_portfolio_risk(weights,stock_data):
    cov_matrix = stock_data.cov()
    corr_matrix = stock_data.corr()
    weights=np.array(weights)
    portfolio_variance = np.dot(weights.T,(np.dot(cov_matrix, weights)))
    portfolio_stddev = np.sqrt(portfolio_variance)*100*np.sqrt(250)
    return(portfolio_stddev,corr_matrix)

def calculate_sharpe_ratio(weights,avg_returns,cov_matrix):
    global risk_free_rate_per_cent
    portfolio_return = np.dot(avg_returns, weights)*100
    portfolio_variance = (np.dot(weights.T, np.dot(cov_matrix, weights))) * 250
    portfolio_stddev = np.sqrt(portfolio_variance) * 100
    sharpe = (portfolio_return - risk_free_rate_per_cent) / portfolio_stddev
    return (-sharpe)

def linear_regression(x,y):
    global daily_risk_free_rate
    y=y-daily_risk_free_rate
    x=x-daily_risk_free_rate
    slope, intercept, r_value, p_value, std_err = stats.linregress(x, y)
    r = r_value ** 2

    cov = pd.DataFrame(np.cov(x,y) * 250)
    market_var = x.var() * 250
    cov_with_market = cov.iloc[0, 1]

    beta = cov_with_market / market_var

    return(beta,r)

def calculate_alpha(portfolio_annualised_return_per_cent,sp_annualised_return_per_cent,beta):
    global risk_free_rate_per_cent
    capm_value=risk_free_rate_per_cent+beta*(sp_annualised_return_per_cent-risk_free_rate_per_cent)
    alpha=portfolio_annualised_return_per_cent-capm_value
    print(capm_value,alpha)
    return(alpha,capm_value)

def saving_1(portfolio,results,correlation):
    global path
    global date_index
    portfolio_df = pd.DataFrame()
    portfolio_df['Date'] = date_index
    portfolio_df['Portfolio daily return %'] = portfolio['portfolio']
    dfs = {'correlation matrix': correlation, 'analysis': results,}
    writer = pd.ExcelWriter(path + 'Portfolio_Analysis.xlsx', engine='xlsxwriter')
    for sheet_name in dfs.keys():
        if sheet_name == 'analysis':
            dfs[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)
        else:
            dfs[sheet_name].to_excel(writer, sheet_name=sheet_name, index=True)
        for i, col in enumerate(dfs[sheet_name].columns):
            worksheet = writer.sheets[sheet_name]
            width = max(dfs[sheet_name][col].apply(lambda x: len(str(x))).max(), len(dfs[sheet_name][col]))
            worksheet.set_column(i, i, width)
    portfolio_df.to_excel(writer, sheet_name='portfolio daily returns', index=False)
    for i, col in enumerate(portfolio_df.columns):
        worksheet = writer.sheets['portfolio daily returns']
        width = max(portfolio_df[col].apply(lambda x: len(str(x))).max(), len(col))
        worksheet.set_column(i, i, width)
    writer.close()

def saving_2(portfolio,results,correlation,optimal_weights,stocks):
    global path
    global date_index
    portfolio_df=pd.DataFrame()
    portfolio_df['Date']=date_index
    portfolio_df['Portfolio daily return %']=portfolio['portfolio']
    optimal_weights=pd.DataFrame(optimal_weights)
    optimal_weights.index=stocks
    dfs={'correlation matrix':correlation, 'analysis':results,'optimal weights':optimal_weights}
    writer=pd.ExcelWriter(path+'Portfolio_Analysis.xlsx',engine='xlsxwriter')
    for sheet_name in dfs.keys():
        if sheet_name == 'analysis':
            dfs[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)
        else:
            dfs[sheet_name].to_excel(writer, sheet_name=sheet_name, index=True)
        for i, col in enumerate(dfs[sheet_name].columns):
            worksheet = writer.sheets[sheet_name]
            width = max(dfs[sheet_name][col].apply(lambda x: len(str(x))).max(), len(dfs[sheet_name][col]))
            worksheet.set_column(i, i, width)
    portfolio_df.to_excel(writer, sheet_name='portfolio daily returns', index=False)
    for i, col in enumerate(portfolio_df.columns):
        worksheet = writer.sheets['portfolio daily returns']
        width = max(portfolio_df[col].apply(lambda x: len(str(x))).max(), len(col))
        worksheet.set_column(i, i, width)
    writer.close()

def updating():
    label1.place(x=150, y=80, width=200, height=20)
    label1.config(text='Done')

window=Tk()
window.geometry('500x500')
window.title('Portfolio Analysis and Optimization')

label1=Label(window,text='Ticker')
label2=Label(window,text='weight(eg: msft 0.2)')
entry1=Entry(window)
entry2=Entry(window)
entry3=Entry(window)
entry4=Entry(window)
entry5=Entry(window)
entry6=Entry(window)
button1=Button(window,text='Portfolio Analysis',command=mode_1)
button2=Button(window,text='Portfolio Optimization',command=mode_2)

objects=[label1,label2,entry1,entry2,entry3,entry4,entry5,entry6,button1,button2]

button1.place(x=20, y=80, width=200, height=20)
button2.place(x=240, y=80, width=200, height=20)

#INITIALIZE

global path
path=" "
global risk_free_rate_per_cent
risk_free_rate_per_cent=2
global inflation_rate
inflation_rate=2
global n
n=5
global daily_risk_free_rate
daily_risk_free_rate = ((1 + risk_free_rate_per_cent/100) ** (1 / 250) - 1)
daily_risk_free_rate_per_cent=daily_risk_free_rate*100
daily_risk_free_rate_per_cent='{:f}'.format(daily_risk_free_rate_per_cent)
global mode
mode=1


window.mainloop()