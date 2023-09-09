import os
import time
import yfinance as yf
import numpy as np
import pandas as pd
from io import StringIO
from datetime import datetime
import selenium.common.exceptions
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from rich.console import Console


console = Console()


class StocksExplorer:
    def __init__(self) -> None:
        self.tickers = pd.read_csv('Data/Tickers.csv', sep=';')['TICKER']
        self.driver = webdriver.Chrome(service=Service(
            executable_path=ChromeDriverManager().install()))
        self.driver.maximize_window()

        with console.status(status='Coletando dados do site Investidor10...'):
            df = self.coletar_dados()

        with console.status(status='Filtrando dados...'):
            df = self.filtrar_dados(df=df)

        with console.status(status='Calculando indicadores técnicos...'):
            df = self.indicadores_tecnicos(df=df)

        df = df[['Cotação', 'Volat. Anualizada',
                 'Volat. Mensal', 'Cres. LPA (3 Anos)', 'RSI']]
        df.to_excel(
            excel_writer=f"Resultado Ações ({datetime.now().strftime('%d-%m-%Y')}).xlsx", index=True)

    def coletar_dados(self) -> pd.DataFrame:
        qtd_index = 1
        df = pd.DataFrame()
        for ticker in self.tickers:
            self.driver.get(url=f'https://investidor10.com.br/acoes/{ticker}/')
            try:
                table_element = self.driver.find_element(
                    by=By.XPATH, value='//*[@id="table-indicators-history"]')
                time.sleep(1)

                table_html = table_element.get_attribute("outerHTML")

                # Envolve a string HTML em um objeto StringIO
                html_buffer = StringIO(table_html)

                # Use o objeto StringIO com a função read_html
                df_tmp = pd.read_html(html_buffer)[0]

                df_tmp.rename(
                    columns={'Unnamed: 0': 'Indicadores'}, inplace=True)
                df_tmp.set_index('Indicadores', drop=True, inplace=True)
                df_tmp = df_tmp.apply(lambda x: x.map(
                    lambda x: str(x).replace('%', '').replace(',', '')))
                df_tmp = df_tmp.apply(lambda x: x.map(
                    lambda x: 0 if x == "-" else x))
                df_tmp = df_tmp.apply(lambda x: x.map(lambda x: int(x) / 100))

                # Trocar índice e colunas
                df_tmp = df_tmp.transpose()
                df_tmp['Ticker'] = ticker
                df_tmp['Cotação'] = self.driver.find_element(
                    by=By.XPATH, value='//*[@id="cards-ticker"]/div[1]/div[2]/div/span').text
                df = pd.concat([df, df_tmp])
                console.print(
                    f"[[blue]{datetime.now().strftime('%H:%M:%S')}[/]] ---> [[yellow]Dados coletados para o ativo[/]] :: [{ticker}] ---> [[yellow]Empresas[/]] :: [{qtd_index} de {len(self.tickers)}]")

            except IndexError:
                pass

            except ValueError:
                pass

            except selenium.common.exceptions.NoSuchElementException as e:
                console.print(
                    f"[[red]Erro[/]] Não foi possível coletar dados do ativo {ticker} :: {str(e)}")

            qtd_index += 1

        return df

    @staticmethod
    def indicadores_tecnicos(df: pd.DataFrame):
        df_copy = df.copy()
        df_copy['Volat. Anualizada'] = np.nan
        df_copy['Volat. Mensal'] = np.nan
        df_copy['RSI'] = np.nan

        for index, _ in df.iterrows():
            df_yf = yf.Ticker(
                ticker=index + ".SA").history(period='1y', interval='1d')

            if df_yf is not df_yf.empty:
                # Calcular as volatilidades
                desvio_padrao_diario = df_yf['Close'].std()
                # Calcular a volatilidade anualizada
                volatilidade_anualizada = round(
                    desvio_padrao_diario * np.sqrt(252), 2)
                df_copy.loc[index,
                            'Volat. Anualizada'] = volatilidade_anualizada

                volatilidade_mensal = round(
                    desvio_padrao_diario * np.sqrt(21), 2)
                df_copy.loc[index, 'Volat. Mensal'] = volatilidade_mensal

                console.print(
                    f"[[blue]{datetime.now().strftime('%H:%M:%S')}[/]] ---> [[yellow]Volatilidades calculadas[/]] [[yellow]Ticker[/]] :: [{index}] [[yellow]Volat. Anual[/]] :: [{volatilidade_anualizada}] [[yellow]Volat. Mensal[/]] :: [{volatilidade_mensal}]"
                )

                # Calculando RSI
                try:
                    n = 14

                    def rma(x, n, y0):
                        a = (n - 1) / n
                        ak = a**np.arange(len(x) - 1, -1, -1)
                        return np.r_[np.full(n, np.nan), y0, np.cumsum(
                            ak * x) / ak / n + y0 * a**np.arange(1, len(x) + 1)]

                    df_yf['change'] = df_yf['Close'].diff()
                    df_yf['gain'] = df_yf.change.mask(df_yf.change < 0, 0.0)
                    df_yf['loss'] = -df_yf.change.mask(df_yf.change > 0, -0.0)
                    df_yf['avg_gain'] = rma(
                        df_yf.gain[n + 1:].to_numpy(), n, np.nansum(df_yf.gain.to_numpy()[:n + 1]) / n)
                    df_yf['avg_loss'] = rma(
                        df_yf.loss[n + 1:].to_numpy(), n, np.nansum(df_yf.loss.to_numpy()[:n + 1]) / n)
                    df_yf['rs'] = df_yf.avg_gain / df_yf.avg_loss
                    df_yf['rsi'] = 100 - (100 / (1 + df_yf.rs))
                    rsi = round(df_yf['rsi'][-1:].values[0], 2)
                    df_copy.loc[index, 'RSI'] = rsi

                    console.print(
                        f"[[blue]{datetime.now().strftime('%H:%M:%S')}[/]] ---> [[yellow]RSI calculado[/]] [[yellow]Ticker[/]] :: [{index}] [[yellow]RSI[/]] :: [{rsi}]"
                    )

                except ValueError:
                    pass

        return df_copy

    @staticmethod
    def filtrar_dados(df: pd.DataFrame) -> pd.DataFrame:
        atual_year = int(datetime.now().strftime('%Y'))
        df_copy = df.copy()

        filtro_ = \
            (df_copy['P/L'] < 8) & \
            (df_copy['P/VP'] > 0) & \
            (df_copy['P/VP'] < 2) & \
            (df_copy['P/EBIT'] > 0) & \
            (df_copy['P/EBIT'] < 10) & \
            (df_copy['DIVIDEND YIELD (DY)'] > 6) & \
            (df_copy['ROE'] > 15) & \
            (df_copy['ROIC'] > 10) & \
            (df_copy['P/ATIVO'] < 2) & \
            (df_copy['P/RECEITA (PSR)'] < 2) & \
            (df_copy['DÍVIDA LÍQUIDA / EBITDA'] < 3)

        df_filtrado = df_copy[filtro_]
        df_filtrado = df_filtrado[df_filtrado.index ==
                                  'Atual'].reset_index(drop=True)

        for index, row in df_filtrado.iterrows():
            lpa_a = df_copy[(df_copy['Ticker'] == row['Ticker']) & (
                df_copy.index == str(atual_year - 1))]['LPA'].values[0]
            lpa_b = df_copy[(df_copy['Ticker'] == row['Ticker']) & (
                df_copy.index == str(atual_year - 2))]['LPA'].values[0]
            df_filtrado.loc[index,
                            'Cres. LPA (3 Anos)'] = row['LPA'] > lpa_a > lpa_b

        df_filtrado.set_index('Ticker', inplace=True)
        df_filtrado = df_filtrado[df_filtrado['Cres. LPA (3 Anos)'] == True]
        df_filtrado.rename(columns={'Indicadores': ''}, inplace=True)

        return df_filtrado


if __name__ == "__main__":
    app = StocksExplorer()
