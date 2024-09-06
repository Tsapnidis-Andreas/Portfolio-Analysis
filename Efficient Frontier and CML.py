import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import yfinance as yf
import datetime as dt
from datetime import datetime
from scipy.optimize import minimize
import openpyxl
import xlsxwriter
from xlsxwriter import Workbook
def obtaining_returns(stocks):
    for i in stocks:
        end_date = datetime.now().strftime('%Y-%m-%d')
        start_date = (datetime.strptime(end_date, '%Y-%m-%d') - dt.timedelta(days=365 * 5)).strftime('%Y-%m-%d')
        stock_data = yf.download(i, start=start_date, end=end_date)
        stock_data = pd.DataFrame(stock_data)
        data[str(i)] = stock_data['Close']

    log_returns = np.log(data / data.shift(1))
    avg_returns = log_returns.mean() * 250
    cov_matrix = log_returns.cov() * 250
    return(avg_returns,cov_matrix)

def random_portfolios(avg_returns,cov_matrix):
    portfolio_returns = []
    portfolio_volatilities = []

    for i in range(1, 8000):
        weights = np.random.random(len(stocks))
        weights /= np.sum(weights)
        portfolio_returns.append(np.sum(weights * avg_returns))
        portfolio_volatilities.append(np.sqrt(np.dot(weights.T, np.dot(cov_matrix, weights))))

    portfolio_returns = np.array(portfolio_returns)
    portfolio_volatilities = np.array(portfolio_volatilities)

    portfolios = pd.DataFrame({'Return': portfolio_returns, 'Volatility': portfolio_volatilities})
    return(portfolios)
def minimize_volatility(weights):
    weights=np.array(weights)
    v=np.sqrt(np.dot(weights.T,np.dot(cov_matrix,weights)))
    return(v)

def get_return(weights):
    r=np.sum(weights * avg_returns)
    return(r)

def CML(vol):
    global max_ret
    global max_vol
    global risk_free_rate
    slope = (max_ret-risk_free_rate)/max_vol
    constant=risk_free_rate
    r=constant+slope*vol
    return(r)

def plotting(x,y,x1,optimal_portfolios):
    figure = plt.figure(figsize=(18, 10))
    plt.scatter(x, y)
    plt.xlabel('Expected Volatility')
    plt.ylabel('Expected Return')
    plt.plot(optimal_portfolios['volatility %'], optimal_portfolios['return %'], c='g')
    plt.scatter(optimal_portfolios['volatility %'], optimal_portfolios['return %'], c='g')
    plt.plot(x1, CML(x1), c='r')
    plt.title('Efficient Frontier')
    plt.savefig(path + 'efficient frontier.png')

def saving(optimal_portfolios):
    writer = pd.ExcelWriter(path + 'Optimal Portfolios.xlsx', engine='xlsxwriter')
    optimal_portfolios.to_excel(writer, index=False, header=True, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    for i, col in enumerate(optimal_portfolios.columns):
        width = max(optimal_portfolios[col].apply(lambda x: len(str(x))).max(), len(optimal_portfolios[col]))
        worksheet.set_column(i, i, width)
    writer.close()


#INITIALIZE
data=pd.DataFrame()
stocks=['msft','jpm','ko']
global risk_free_rate
risk_free_rate=2
global path
path=" "

avg_returns,cov_matrix=obtaining_returns((stocks))

portfolios=random_portfolios(avg_returns,cov_matrix)


optimal_portfolios=pd.DataFrame()
returns=np.linspace(portfolios['Return'].min(),portfolios['Return'].max(),30)
optimal_portfolios['return %']=(returns)
optimal_portfolios['volatility %']=np.ones((len(optimal_portfolios),1))

weights_df=pd.DataFrame()

volatility_opt=[]
bounds = [(0, 1) for i in range(len(stocks))]
initial_weights = np.array([1 / len(stocks)] * len(stocks))
constraints=({'type': 'eq', 'fun': lambda weights: np.sum(weights) - 1},
             {'type':'eq','fun':lambda weights:get_return(weights)-r})

k=0
for r in returns:
    opt=minimize(minimize_volatility,initial_weights,method='SLSQP',bounds=bounds,constraints=constraints)
    volatility_opt.append(opt['fun'])
    optimal_weights =opt.x
    optimal_weights=optimal_weights.tolist()

    for i in range(1,len(stocks)+1):
        a=optimal_weights[i-1]*100
        a='{:f}'.format(a)
        weights_df.loc[k,i]=a
    k=k+1


weights_df.columns=range(1,len(stocks)+1)

optimal_portfolios['return %']=returns*100
optimal_portfolios['volatility %']=volatility_opt
optimal_portfolios['volatility %']=optimal_portfolios['volatility %']*100

for i in range(1,len(stocks)+1):
    optimal_portfolios[stocks[i-1]]=weights_df[i]

optimal_portfolios['sharpe ratio']=(optimal_portfolios['return %']-risk_free_rate)/optimal_portfolios['volatility %']

col='sharpe ratio'
max_i=optimal_portfolios.loc[optimal_portfolios[col].idxmax()]
max_i=max_i.to_list()

global max_ret
global max_vol

max_ret=float(max_i[0])
max_vol=float(max_i[1])

x1=np.linspace(0,max_vol*1.2,100)


x=portfolios['Volatility']*100
y=portfolios['Return']*100

plotting(x,y,x1,optimal_portfolios)

saving(optimal_portfolios)