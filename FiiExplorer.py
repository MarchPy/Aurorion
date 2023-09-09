import json
import time
import numpy as np
import pandas as pd
import yfinance as yf
import selenium.common.exceptions
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from rich.console import Console
from datetime import datetime


console = Console()


class FiiExplorer:
    def __init__(self) -> None:
        try:
            with console.status(status='Inicializando driver...'):
                self.driver = webdriver.Chrome(
                    service=Service(ChromeDriverManager().install()))
                self.driver.maximize_window()

            with console.status(status='Coletando os tickers...'):
                df_tickers = self.coletar_tickers()

            with console.status(status='Coletando dados do site Status Invest...'):
                df_final = self.coletar_dados_do_ticker(tickers=df_tickers)

            with console.status(status='Tratando dados coletados...'):
                df_final = self.tratar_dados(df=df_final)

            with console.status(status='Filtrando dados...'):
                df_final = self.filtro(df=df_final)

            with console.status(status='Calculando indicadores técnicos...'):
                df_final = self.indicadores_tecnicos(df=df_final)

            with console.status(status='Calculando AHP Gaussiano...'):
                df_final = self.ahp_gaussian(df=df_final)

            with console.status(status='Salvando dados em formato de excel...'):
                df_final = df_final[['Nome', 'Setor', 'Tipo', 'Cotação Atual', 'DY (12 Meses)',
                                     'Volat. Anualizada', 'Volat. Mensal', 'RSI', 'Ranking AHP']]
                df_final.to_excel(
                    excel_writer=f"Resultado FIIs ({datetime.now().strftime('%d-%m-%Y')}).xlsx", index=True)

        except KeyboardInterrupt:
            console.print(
                f"[[blue]{datetime.now().strftime('%H:%M:%S')}[/]] ---> [[green]Processo encerrado pelo usuário[/]]")

    def coletar_dados_do_ticker(self, tickers) -> pd.DataFrame:
        qtd_index = 1
        df = pd.DataFrame()

        for _, row in tickers.iterrows():
            url = "https://statusinvest.com.br/fundos-imobiliarios/" if row[
                'Tipo'] == 'Fundo Imobiliário' else 'https://statusinvest.com.br/fiagros/'
            url = url + row['Ticker']
            self.driver.get(url=url)

            time.sleep(1)
            try:
                if row['Tipo'] == 'Fundo Imobiliário':
                    nome_ativo = self.driver.find_element(
                        by=By.XPATH, value='//*[@id="main-header"]/div[2]/div/div[1]/h1/small').text
                    cotacao_atual = self.driver.find_element(
                        by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[1]/div[1]/div/div[1]/strong').text
                    cotacao_minima52 = self.driver.find_element(
                        by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[1]/div[2]/div/div[1]/strong').text
                    cotacao_maxima52 = self.driver.find_element(
                        by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[1]/div[3]/div/div[1]/strong').text
                    dy_ultimo_12_meses = self.driver.find_element(
                        by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[1]/div[4]/div/div[1]/strong').text
                    valorizacao_12_meses = self.driver.find_element(
                        by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[1]/div[5]/div/div[1]/strong').text
                    valorizacao_mes_atual = self.driver.find_element(
                        by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[1]/div[5]/div/div[2]/div/span[2]/b').text
                    pvp = self.driver.find_element(
                        by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[5]/div/div[2]/div/div[1]/strong').text
                    valor_em_caixa = self.driver.find_element(
                        by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[5]/div/div[3]/div/div[1]/strong').text
                    dy_cagr_3_anos = self.driver.find_element(
                        by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[5]/div/div[4]/div/div[1]/strong').text
                    valor_cagr_3_anos = self.driver.find_element(
                        by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[5]/div/div[5]/div/div[1]/strong').text
                    numero_de_cotistas = self.driver.find_element(
                        by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[5]/div/div[6]/div/div[1]/strong').text
                    liquidez_media_diaria = self.driver.find_element(
                        by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[6]/div/div/div[3]/div/div/div/strong').text
                    rendimento_mensal_medio_24_meses = self.driver.find_element(
                        by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[6]/div/div/div[1]/div/div/strong').text
                    ultimo_rendimento = self.driver.find_element(
                        by=By.XPATH, value='//*[@id="dy-info"]/div/div[1]/strong').text
                    segmento = self.driver.find_element(
                        by=By.XPATH, value='//*[@id="fund-section"]/div/div/div[4]/div/div[1]/div/div/div/a/strong').text

                else:
                    nome_ativo = self.driver.find_element(
                        by=By.XPATH, value='//*[@id="main-header"]/div[2]/div/div[1]/h1/small').text
                    cotacao_atual = self.driver.find_element(
                        by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[1]/div[1]/div/div[1]/strong').text
                    cotacao_minima52 = self.driver.find_element(
                        by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[1]/div[2]/div/div[1]/strong').text
                    cotacao_maxima52 = self.driver.find_element(
                        by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[1]/div[3]/div/div[1]/strong').text
                    dy_ultimo_12_meses = self.driver.find_element(
                        by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[1]/div[4]/div/div[1]/strong').text
                    valorizacao_12_meses = self.driver.find_element(
                        by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[1]/div[5]/div/div[1]/strong').text
                    valorizacao_mes_atual = self.driver.find_element(
                        by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[1]/div[5]/div/div[2]/div/span[2]/b').text
                    pvp = self.driver.find_element(
                        by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[4]/div/div[2]/div/div[1]/strong').text
                    valor_em_caixa = self.driver.find_element(
                        by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[4]/div/div[3]/div/div[1]/strong').text
                    dy_cagr_3_anos = self.driver.find_element(
                        by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[4]/div/div[4]/div/div[1]/strong').text
                    valor_cagr_3_anos = self.driver.find_element(
                        by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[4]/div/div[5]/div/div[1]/strong').text
                    numero_de_cotistas = self.driver.find_element(
                        by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[4]/div/div[6]/div/div[1]/strong').text
                    liquidez_media_diaria = self.driver.find_element(
                        by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[5]/div/div/div[3]/div/div/div/strong').text
                    rendimento_mensal_medio_24_meses = self.driver.find_element(
                        by=By.XPATH, value='//*[@id="main-2"]/div[2]/div[5]/div/div/div[1]/div/div/strong').text
                    ultimo_rendimento = self.driver.find_element(
                        by=By.XPATH, value='//*[@id="dy-info"]/div/div[1]/strong').text
                    segmento = self.driver.find_element(
                        by=By.XPATH, value='//*[@id="fund-section"]/div/div/div[4]/div/div[1]/div/div/div/strong').text

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
                console.print(
                    f"[[blue]{datetime.now().strftime('%H:%M:%S')}[/]] ---> [[yellow]Dados coletados para o ativo[/]] :: [{row['Ticker']}] ---> [[yellow]Empresas[/]] :: [{qtd_index} de {len(tickers)}]")

            except selenium.common.exceptions.NoSuchElementException as e:
                console.print(
                    f"[[red]Erro[/]] Não foi possível coletar dados do ativo {row['Ticker']} :: {str(e)}")

            except KeyboardInterrupt:
                break

            qtd_index += 1

        self.driver.close()

        return df

    @staticmethod
    def coletar_tickers() -> pd.DataFrame:
        tickers = pd.read_csv(
            filepath_or_buffer='Data/TickersFIIs.csv', encoding='iso-8859-1', sep=';')
        
        return tickers

    @staticmethod
    def tratar_dados(df: pd.DataFrame) -> pd.DataFrame:
        df = df.apply(lambda x: x.map(lambda y: y.replace('%', '')))
        df = df.apply(lambda x: x.map(lambda y: y.replace('.', '').replace(',', '.')))
        df = df.apply(lambda x: x.map(lambda y: 0 if x == "-" else y))

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

    @staticmethod
    def filtro(df: pd.DataFrame) -> pd.DataFrame:
        with open(file='Configs/Fiis_config.json', mode='r', encoding='utf-8') as file_obj:
            config = json.load(file_obj)

        param_filtro_ = \
            (df['DY (12 Meses)'] >= config['DY (Min)']) & \
            (df['P/VP'] >= config['PVP (Min)']) & \
            (df['P/VP'] <= config['PVP (Max)']) & \
            (df['N° de cotistas'] >= config['N de cotistas (Min)']) & \
            (df['Liquidez média diária'] >= config['Liquidez média diária (Min)']) & \
            (df['Valorização (12 Meses)'] >= config['Valorização (12 Meses) (Min)']) & \
            (df['Valorização (Mês atual)'] >=
             config['Valorização (Mês atual) (Min)'])

        df = df[param_filtro_]

        return df

    @staticmethod
    def ahp_gaussian(df: pd.DataFrame) -> pd.DataFrame:
        df_copy = df.copy()
        df_final = df.copy()

        df_copy = df_copy[[
            'DY (12 Meses)', 'P/VP', 'N° de cotistas', 'DY CAGR (3 Anos)',
            'Liquidez média diária', 'Rendimento mensal médio (24 Meses)',
            'Valorização (12 Meses)', 'Valorização (Mês atual)', 'Valor em Caixa',
            'Volat. Anualizada', 'Volat. Mensal'
        ]]

        crit_max_min = {
            'DY (12 Meses)': 'MAX',
            'P/VP': 'MIN',
            'N° de cotistas': 'MAX',
            'DY CAGR (3 Anos)': 'MAX',
            'Liquidez média diária': 'MAX',
            'Rendimento mensal médio (24 Meses)': 'MAX',
            'Valorização (12 Meses)': 'MAX',
            'Valorização (Mês atual)': 'MAX',
            'Valor em Caixa': 'MIN',
            'Volat. Anualizada': 'MAX',
            'Volat. Mensal': 'MAX',
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

        # Adicionar as colunas de média, desvio padrão, Fator Gaussiano, sensibilidade e ranking final ao DataFrame
        df_copy['Media'] = media
        df_copy['Desvio Padrao'] = desvio_padrao
        df_copy['Fator Gaussiano'] = fator_gaussiano
        df_copy['Analise Sensibilidade'] = sensibilidade
        df_copy['Ranking Final'] = rank_final

        for index, _ in df_final.iterrows():
            df_final.loc[index, 'Ranking AHP'] = df_copy.loc[index, 'Ranking Final']

        df_final.sort_values(by='Ranking AHP', ascending=True, inplace=True)
        return df_final

    @staticmethod
    def definir_quantidade_de_compra(df: pd.DataFrame, ammount: int = 1000) -> pd.DataFrame:
        df_copy = df.copy()
        parts = [
            60, 10, 10, 10, 10, 10
        ]

        i = 0
        for index, row in df.iterrows():
            if i == len(parts) - 1:
                break

            qtd = int((ammount * (parts[i] / 100)) / row['Cotação Atual'])
            df.loc[index, 'Quant. de compra'] = qtd
            i += 1

        return df_copy

    @staticmethod
    def indicadores_tecnicos(df: pd.DataFrame) -> pd.DataFrame:
        df_copy = df.copy()
        df_copy['Volat. Anualizada'] = np.nan
        df_copy['Volat. Mensal'] = np.nan
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
                df_copy.loc[index, 'Volat. Anualizada'] = volatilidade_anualizada

                volatilidade_mensal = round(
                    desvio_padrao_diario * np.sqrt(21), 2)
                df_copy.loc[index, 'Volat. Mensal'] = volatilidade_mensal

                # Calculando RSI
                try:
                    period = 14

                    def rma(x, n, y0):
                        a = (n-1) / n
                        ak = a**np.arange(len(x)-1, -1, -1)
                        return np.r_[np.full(n, np.nan), y0, np.cumsum(ak * x) / ak / n + y0 * a**np.arange(1, len(x)+1)]

                    df_yf['change'] = df_yf['Close'].diff()
                    df_yf['gain'] = df_yf.change.mask(df_yf.change < 0, 0.0)
                    df_yf['loss'] = -df_yf.change.mask(df_yf.change > 0, -0.0)
                    df_yf['avg_gain'] = rma(df_yf.gain[period+1:].to_numpy(), period, np.nansum(df_yf.gain.to_numpy()[:period+1])/period)
                    df_yf['avg_loss'] = rma(df_yf.loss[period+1:].to_numpy(), period, np.nansum(df_yf.loss.to_numpy()[:period+1])/period)
                    df_yf['rs'] = df_yf.avg_gain / df_yf.avg_loss
                    df_yf['rsi'] = 100 - (100 / (1 + df_yf.rs))
                    rsi = round(df_yf['rsi'][-1:].values[0], 2)
                    df_copy.loc[index, 'RSI'] = rsi

                    console.print(
                        f"[[blue]{datetime.now().strftime('%H:%M:%S')}[/]] ---> [[yellow]Volatilidades calculadas[/]] [[yellow]Ticker[/]] :: [{index}] [[yellow]Volat. Anual[/]] :: [{volatilidade_anualizada}] [[yellow]Volat. Mensal[/]] :: [{volatilidade_mensal}] [[yellow]RSI[/]] :: [{rsi}]"
                    )

                except ValueError:
                    pass

        return df_copy  


if __name__ == "__main__":
    app = FiiExplorer()
