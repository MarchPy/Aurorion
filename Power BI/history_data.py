import os
import pandas as pd
import yfinance as ynf
from rich.console import Console


console = Console()
save_path = "Power BI/Tickers Data/"


try:
    os.mkdir(path=save_path)
    
except FileExistsError:
    pass


tickers = pd.read_csv(filepath_or_buffer='data/tickers.csv', sep=';', encoding='iso-8859-1')['Ticker'].to_list()
for ticker in tickers:
    try:
        df_yf = ynf.Ticker(ticker=ticker + ".sa").history(period='1y', interval='1d')
        df_yf.reset_index(inplace=True)
        df_yf['Ticker'] = ticker
        df_yf[['Open', 'High', 'Low', 'Close']] = df_yf[['Open', 'High', 'Low', 'Close']].round(2)
        df_yf = df_yf[['Ticker', 'Date', 'Open', 'High', 'Low', 'Close', 'Volume']]
        df_yf.to_csv(path_or_buf=os.path.join(save_path, f"{ticker}.csv"), sep=';', decimal=',')
        console.print(f"[[bold green]Dados coletados do ativo[/]] :: {ticker}")
    
    except ValueError:
        pass