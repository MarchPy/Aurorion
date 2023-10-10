import yfinance as yf
import pandas as pd

index = 1
short_period = 8
long_period = 21
volume_period = 21

tickers = pd.read_csv("data/tickers.csv", sep=';', encoding='iso-8859-1')
tickers = tickers[tickers['Tipo'] == "Acao"]
tickers = tickers['Ticker']


data_list = []
for ticker in tickers:
    print(f"[{ticker}] :: [{index} de {len(tickers)}]")
    
    # Baixa os dados históricos do ativo
    try:
        data = yf.Ticker(ticker=ticker + ".sa").history(period="1y", interval='1d')

        data[f'EMA_{short_period}'] = data['Close'].ewm(span=short_period, adjust=False).mean()
        data[f'EMA_{long_period}'] = data['Close'].ewm(span=long_period, adjust=False).mean()
        data[f'Volume_EMA_{volume_period}'] = data['Volume'].ewm(span=volume_period, adjust=False).mean()

        # Identifica os cruzamentos
        data['Signal'] = 0
        data.loc[data[f'EMA_{short_period}'] > data[f'EMA_{long_period}'], 'Signal'] = 1  # Cruzamento para cima
        data.loc[data[f'EMA_{short_period}'] < data[f'EMA_{long_period}'], 'Signal'] = -1  # Cruzamento para baixo

        # Verifica se houve cruzamento
        if data['Signal'].iloc[-1] != data['Signal'].iloc[-2]: # Verificar se o sinal se inverteu
            if data['Signal'].iloc[-1] == 1: # Verifica se o "Sinal" cruzou para cima 
                if data['Volume'].iloc[-1] > data[f'Volume_EMA_{volume_period}'].iloc[-1] * 2: # Verifica se o volume é maior que duas vezes a média
                    data_list.append([ticker, 'Cima', 'Confirmado'])
                
            elif data['Signal'].iloc[-1] == -1: # Verifica se o "Sinal" cruzou para baixo 
                if data['Volume'].iloc[-1] > data[f'Volume_EMA_{volume_period}'].iloc[-1] * 2: # Verifica se o volume é maior que duas vezes a média
                    data_list.append([ticker, 'Baixo', 'Confirmado'])

    except (ValueError, IndexError):
        pass
    
    index += 1
    

df_final = pd.DataFrame(data=data_list, columns=['Ticker', 'Cruzamento', 'Volume'])
df_final.to_excel("Teste.xlsx")
print(df_final)
