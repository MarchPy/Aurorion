import os
import json
import time
import numpy as np
import pandas as pd
import yfinance as yf
import selenium.common.exceptions
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from rich.console import Console


class FundExplorer:
    def __init__(self) -> None:
        self.__console = Console()

    @staticmethod
    def coletar_tickers() -> pd.DataFrame:
        tickers = pd.read_csv(filepath_or_buffer='data/tickers.csv', encoding='iso-8859-1', sep=';').sort_values(by='Ticker')

        return tickers

    def coletar_dados_do_ticker(self, tickers: pd.DataFrame, verbose: bool = False) -> pd.DataFrame:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
        driver.maximize_window()
        qtd_index = 1
        df = pd.DataFrame()

        for _, row in tickers.iterrows():
            url = "https://statusinvest.com.br/fundos-imobiliarios/" if row['Tipo'] == 'Fundo Imobiliário' else 'https://statusinvest.com.br/fiagros/'
            url = url + row['Ticker']
            driver.get(url=url)

            time.sleep(1)
            try:
                if row['Tipo'] == 'Fundo Imobiliário':
                    nome_ativo                       = driver.find_element( by=By.XPATH, value='//*[@id="main-header"]/div[2]/div/div[1]/h1/small').text
                    cotacao_atual                    = driver.find_element(by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[1]/div[1]/div/div[1]/strong').text
                    cotacao_minima52                 = driver.find_element(by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[1]/div[2]/div/div[1]/strong').text
                    cotacao_maxima52                 = driver.find_element(by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[1]/div[3]/div/div[1]/strong').text
                    dy_ultimo_12_meses               = driver.find_element(by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[1]/div[4]/div/div[1]/strong').text
                    valorizacao_12_meses             = driver.find_element(by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[1]/div[5]/div/div[1]/strong').text
                    valorizacao_mes_atual            = driver.find_element(by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[1]/div[5]/div/div[2]/div/span[2]/b').text
                    pvp                              = driver.find_element(by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[5]/div/div[2]/div/div[1]/strong').text
                    valor_em_caixa                   = driver.find_element(by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[5]/div/div[3]/div/div[1]/strong').text
                    dy_cagr_3_anos                   = driver.find_element(by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[5]/div/div[4]/div/div[1]/strong').text
                    valor_cagr_3_anos                = driver.find_element(by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[5]/div/div[5]/div/div[1]/strong').text
                    numero_de_cotistas               = driver.find_element(by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[5]/div/div[6]/div/div[1]/strong').text
                    liquidez_media_diaria            = driver.find_element(by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[6]/div/div/div[3]/div/div/div/strong').text
                    rendimento_mensal_medio_24_meses = driver.find_element(by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[6]/div/div/div[1]/div/div/strong').text
                    ultimo_rendimento                = driver.find_element(by=By.XPATH, value='//*[@id="dy-info"]/div/div[1]/strong').text
                    segmento                         = driver.find_element(by=By.XPATH, value='//*[@id="fund-section"]/div/div/div[4]/div/div[1]/div/div/div/a/strong').text

                else:
                    nome_ativo                       = driver.find_element(by=By.XPATH, value='//*[@id="main-header"]/div[2]/div/div[1]/h1/small').text
                    cotacao_atual                    = driver.find_element(by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[1]/div[1]/div/div[1]/strong').text
                    cotacao_minima52                 = driver.find_element(by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[1]/div[2]/div/div[1]/strong').text
                    cotacao_maxima52                 = driver.find_element(by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[1]/div[3]/div/div[1]/strong').text
                    dy_ultimo_12_meses               = driver.find_element(by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[1]/div[4]/div/div[1]/strong').text
                    valorizacao_12_meses             = driver.find_element(by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[1]/div[5]/div/div[1]/strong').text
                    valorizacao_mes_atual            = driver.find_element(by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[1]/div[5]/div/div[2]/div/span[2]/b').text
                    pvp                              = driver.find_element(by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[4]/div/div[2]/div/div[1]/strong').text
                    valor_em_caixa                   = driver.find_element(by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[4]/div/div[3]/div/div[1]/strong').text
                    dy_cagr_3_anos                   = driver.find_element(by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[4]/div/div[4]/div/div[1]/strong').text
                    valor_cagr_3_anos                = driver.find_element(by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[4]/div/div[5]/div/div[1]/strong').text
                    numero_de_cotistas               = driver.find_element(by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[4]/div/div[6]/div/div[1]/strong').text
                    liquidez_media_diaria            = driver.find_element(by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[5]/div/div/div[3]/div/div/div/strong').text
                    rendimento_mensal_medio_24_meses = driver.find_element(by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[5]/div/div/div[1]/div/div/strong').text
                    ultimo_rendimento                = driver.find_element(by=By.XPATH, value='//*[@id="dy-info"]/div/div[1]/strong').text
                    segmento                         = driver.find_element(by=By.XPATH, value='//*[@id="fund-section"]/div/div/div[4]/div/div[1]/div/div/div/strong').text

                line = {
                    "Ticker": [row['Ticker']],
                    "Nome": [nome_ativo],
                    "Setor": [segmento],
                    "Tipo": [row['Tipo']],
                    "Cotação Atual": [cotacao_atual],
                    "Cotação mínima (52 Meses)": [cotacao_minima52],
                    "Cotação máxima (52 Meses)": [cotacao_maxima52],
                    "DY (12 Meses)": [dy_ultimo_12_meses],
                    "Valorização (12 Meses)": [valorizacao_12_meses],
                    "Valorização (Mês atual)": [valorizacao_mes_atual],
                    "P/VP": [pvp],
                    "Valor em Caixa": [valor_em_caixa],
                    "DY CAGR (3 Anos)": [dy_cagr_3_anos],
                    "Valor CAGR (3 Anos)": [valor_cagr_3_anos],
                    "N° de cotistas": [numero_de_cotistas],
                    "Liquidez média diária": [liquidez_media_diaria],
                    "Rendimento mensal médio (24 Meses)": [rendimento_mensal_medio_24_meses],
                    "Último rendimento": [ultimo_rendimento],
                    "Link": [url]
                }

                df_tmp = pd.DataFrame(line)
                df = pd.concat([df, df_tmp])
                if verbose is True:
                    self.__console.print(
                        f"[[blue]{datetime.now().strftime('%H:%M:%S')}[/]] ---> [[yellow]Dados coletados para o ativo[/]] :: [{row['Ticker']}] [[yellow]Empresas[/]] :: [{qtd_index} de {len(tickers)}]")

            except selenium.common.exceptions.NoSuchElementException as e:
                if verbose is True:
                    self.__console.print(
                        f"[[red]Erro[/]] Não foi possível coletar dados do ativo {row['Ticker']} :: {str(e)}")

            qtd_index += 1
            
        driver.close()
        
        return df

    @staticmethod
    def tratar_dados(df: pd.DataFrame) -> pd.DataFrame:
        df = df.apply(lambda x: x.map(lambda x: x.replace('%', '')))
        df = df.apply(lambda x: x.map(lambda x: x.replace('.', '').replace(',', '.')))
        df = df.apply(lambda x: x.map(lambda x: 0 if x == "-" else x))

        columns_to_float = [
            'Cotação Atual', 'Cotação mínima (52 Meses)', 'Cotação máxima (52 Meses)', 'DY (12 Meses)', 'Valorização (12 Meses)',
            'Valorização (Mês atual)', 'P/VP', 'DY CAGR (3 Anos)', 'Valor CAGR (3 Anos)', 'Valor em Caixa',
            'N° de cotistas', 'Liquidez média diária', 'Rendimento mensal médio (24 Meses)', 'Último rendimento'
        ]

        for column_float in columns_to_float:
            df[column_float] = df[column_float].astype(float)
            df[column_float] = df[column_float].apply(lambda x: round(x, 2))

        df.set_index(keys='Ticker', drop=True, inplace=True)

        return df

    def indicadores_tecnicos(self, df: pd.DataFrame, verbose: bool = False) -> pd.DataFrame:
        df_copy = df.copy()
        df_copy['% Volat. Anualizada'] = np.nan
        df_copy['% Volat. Mensal'] = np.nan
        df_copy['RSI'] = np.nan

        for index, _ in df.iterrows():
            df_yf = yf.Ticker(
                ticker=str(index) + ".SA").history(period='1y', interval='1d')

            if df_yf is not df_yf.empty and len(df_yf) >= 250:
                # Calcular as volatilidades
                desvio_padrao_diario = df_yf['Close'].std()
                # Calcular a volatilidade anualizada
                volatilidade_anualizada = round(
                    desvio_padrao_diario * np.sqrt(252), 2)
                df_copy.loc[index, '% Volat. Anualizada'] = volatilidade_anualizada

                volatilidade_mensal = round(
                    desvio_padrao_diario * np.sqrt(21), 2)
                df_copy.loc[index, '% Volat. Mensal'] = volatilidade_mensal

                # Calculando RSI
                try:
                    period = 14

                    def rma(x, n, y0):
                        a = (n - 1) / n
                        ak = a**np.arange(len(x) - 1, -1, -1)
                        return np.r_[np.full(n, np.nan), y0, np.cumsum(
                            ak * x) / ak / n + y0 * a**np.arange(1, len(x) + 1)]

                    df_yf['change'] = df_yf['Close'].diff()
                    df_yf['gain'] = df_yf.change.mask(df_yf.change < 0, 0.0)
                    df_yf['loss'] = -df_yf.change.mask(df_yf.change > 0, -0.0)
                    df_yf['avg_gain'] = rma(df_yf.gain[period + 1:].to_numpy(), period, np.nansum(df_yf.gain.to_numpy()[:period + 1]) / period)
                    df_yf['avg_loss'] = rma(df_yf.loss[period + 1:].to_numpy(), period, np.nansum(df_yf.loss.to_numpy()[:period + 1]) / period)
                    df_yf['rs'] = df_yf.avg_gain / df_yf.avg_loss
                    df_yf['rsi'] = 100 - (100 / (1 + df_yf.rs))
                    rsi = round(df_yf['rsi'][-1:].values[0], 2)
                    df_copy.loc[index, 'RSI'] = rsi

                except ValueError:
                    pass
                
                # Calcule o retorno simples
                preco_inicial = df_yf['Close'].iloc[0]
                preco_final = df_yf['Close'].iloc[-1]
                retorno_simples = ((preco_final - preco_inicial) / preco_inicial) * 100
                retorno_simples = round(retorno_simples, 2)
                df_copy.loc[index, '% Valorização'] = retorno_simples
                
                if verbose is True:
                        self.__console.print(
                            f"[[blue]{datetime.now().strftime('%H:%M:%S')}[/]] ---> [[yellow]Indicadores calculados[/]] ---> [[yellow]Ticker[/]] :: [{index}] [[yellow]Volat. Anual[/]] :: [{volatilidade_anualizada}] [[yellow]Volat. Mensal[/]] :: [{volatilidade_mensal}] [[yellow]RSI[/]] :: [{rsi}] [[yellow]Retorno[/]] :: [{retorno_simples}]"
                        )

        return df_copy

    @staticmethod
    def filtro(df: pd.DataFrame) -> pd.DataFrame:
        with open(file='config/FundConfig.json', mode='r', encoding='utf-8') as file_obj:
            config = json.load(file_obj)
    
            filter_ = \
                (df['DY (12 Meses)'] >= config['DY (Min)']) & \
                (df['P/VP'] >= config['PVP (Min)']) & \
                (df['P/VP'] <= config['PVP (Max)']) & \
                (df['N° de cotistas'] >= config['N de cotistas (Min)']  ) & \
                (df['Liquidez média diária'] >= config['Liquidez média diária (Min)']) & \
                (df['Valorização (12 Meses)'] >= config['Valorização (12 Meses) (Min)']) & \
                (df['% Volat. Anualizada'] <= config['Volat. Anual (Max)'] ) & \
                (df['% Volat. Mensal'] <= config['Volat. Mensal (Max)'])
     
        df = df[filter_]

        return df.reset_index()

    @staticmethod
    def ahp_gaussian(df: pd.DataFrame) -> pd.DataFrame:
        df_copy = df.copy()
        df_final = df.copy()
        df_copy['Ranking Final'] = 0

        df_copy = df_copy[[
            'DY (12 Meses)', 'P/VP',
            'Liquidez média diária', 'Rendimento mensal médio (24 Meses)',
            'Valorização (12 Meses)', 'Valorização (Mês atual)', '% Volat. Mensal',
            '% Volat. Anualizada', 'Valor em Caixa'
        ]]

        crit_max_min = { 
            'DY (12 Meses)': 'MAX',
            'Liquidez média diária': 'MAX',
            'P/VP': 'MIN',
            'Valor em Caixa': 'MIN',
            '% Volat. Mensal': 'MIN',
            '% Volat. Anualizada': 'MIN',

        }

        # Calcular a matriz normalizada
        for col in crit_max_min:
            if crit_max_min[col] == 'MAX':
                df_copy[col] = df_copy[col] / df_copy[col].sum()
            else:
                df_copy[col] = 1 / df_copy[col] / (1 / df_copy[col]).sum()

        # Calcular a média, desvio padrão e o Fator Gaussiano
        media = df_copy.mean()
        desvio_padrao = df_copy.std()
        fator_gaussiano = desvio_padrao / media

        # Calcular a análise de sensibilidade do ranking
        weights = fator_gaussiano.values
        sensibilidade = df_copy.dot(weights)

        # Calcular o ranking final
        rank_final = sensibilidade.rank(ascending=False)

        # Adicionar as colunas de média, desvio padrão, Fator Gaussiano,
        # sensibilidade e ranking final ao DataFrame
        df_copy['Media'] = media
        df_copy['Desvio Padrao'] = desvio_padrao
        df_copy['Fator Gaussiano'] = fator_gaussiano
        df_copy['Analise Sensibilidade'] = sensibilidade
        df_copy['Ranking Final'] = rank_final

        for index, _ in df_final.iterrows():
            df_final.loc[index, 'Ranking AHP'] = df_copy.loc[index, 'Ranking Final']
        
        df_final.sort_values(by='Ranking AHP', ascending=True, inplace=True)

        return df_final.reset_index()

    @staticmethod
    def definir_quantidade_de_compra(df: pd.DataFrame, ammount: int = 220) -> pd.DataFrame:
        df_copy = df.copy()
        parts = [
            40, 20, 10, 10, 10
        ]

        i = 0
        for index, row in df.iterrows():
            if i == len(parts):
                break

            qtd = int((ammount * (parts[i] / 100)) / row['Cotação Atual'])
            df_copy.loc[index, 'Quant. de compra'] = qtd
            i += 1

        return df_copy

    def verificar_investimentos(self, df: pd.DataFrame, filename: str) -> None:
        df_carteira = pd.read_csv(filepath_or_buffer=filename, sep=';')
        
        for _, row in df_carteira.iterrows():
            if row['Ticker'] not in df['Ticker'].values:
                self.__console.print(f"[[blue]{datetime.now().strftime('%H:%M:%S')}[/]] ---> [[yellow]Atenção[/]] ---> [[red]Ticker deve ser retirado da sua carteira[/]] :: [{row['Ticker']}]")

        for ticker_rec in df['Ticker'].values:
            if ticker_rec not in df_carteira.values:
                self.__console.print(f"[[blue]{datetime.now().strftime('%H:%M:%S')}[/]] ---> [[yellow]Atenção[/]] ---> [[green]Ticker deve ser adicionado na sua carteira[/]] :: [{ticker_rec}]")        


def main():
    app = FundExplorer()
    tickers = app.coletar_tickers()
    df_data = app.coletar_dados_do_ticker(tickers=tickers, verbose=True)
    df_tratado = app.tratar_dados(df=df_data)
    df_indicadores_tecnicos = app.indicadores_tecnicos(df=df_tratado, verbose=True)
    df_filtrado = app.filtro(df=df_indicadores_tecnicos)
    
    columns = [
        'Ticker', 'Nome', 'Setor', 'Tipo', 'Cotação Atual', 'DY (12 Meses)',
        'N° de cotistas', '% Volat. Anualizada', '% Volat. Mensal', '% Valorização', 'RSI'
    ]
    
    df_final = df_filtrado[columns]
        
    save_file = True
    if save_file is True:
        save_folder = "resultados/"

        try:
            os.mkdir(save_folder)

        except FileExistsError:
            pass

        filename = f"Resultado FIIs ({datetime.now().strftime('%d-%m-%Y')}).xlsx"
        df_final.to_excel(excel_writer=save_folder + filename, index=False)


if __name__ == "__main__":
    main()
