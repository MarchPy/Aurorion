import json
import time
import sqlite3
import pandas as pd
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
        with console.status(status='Inicializando driver...'):
            self.driver = webdriver.Chrome(
                service=Service(ChromeDriverManager().install()))
            self.driver.maximize_window()

        with console.status(status='Coletando os tickers...'):
            df_tickers = self.colect_tickers()

        with console.status(status='Coletando dados do site Status Invest...'):
            df_data = self.colect_data_from_ticker(tickers=df_tickers)

        with console.status(status='Tratando dados coletados...'):
            df_tratado = self.tratar_dados(df=df_data)

        with console.status(status='Filtrando dados...'):
            df_filtrado = self.filtro(df=df_tratado)

        with console.status(status='Inserindo dados no banco de dados...'):
            self.database(df=df_filtrado)

        with console.status(status='Aplicando AHP Gaussiano...'):
            df_final = self.ahp_gaussian(df=df_filtrado)

        with console.status(status='Salvando dados em formato de excel...'):
            df_final = df_final[['Nome', 'Setor', 'Tipo',
                                 'Cotação Atual', 'DY (12 Meses)', 'Ranking AHP Gaussiano']]
            df_final.to_excel(
                excel_writer=f"Resultado FIIs ({datetime.now().strftime('%d-%m-%Y')}).xlsx", index=True)

    @staticmethod
    def database(df: pd.DataFrame):
        with sqlite3.connect(database='Data/History.sqlite') as conn:
            cursor = conn.cursor()
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS History (
                    data TEXT,
                    ticker TEXT,
                    nome TEXT,
                    setor TEXT,
                    dy FLOAT, 
                    pvp FLOAT,
                    liquidez_media_diaria INT
                );
                """
            )

            for index, row in df.iterrows():
                cursor.execute(
                    """
                    INSERT INTO History
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                    """, (datetime.now().strftime('%d/%m/%Y'), index, row['Nome'], row['Setor'], row['DY (12 Meses)'], row['P/VP'], row['Liquidez média diária'])
                )

    @staticmethod
    def colect_tickers() -> pd.DataFrame:
        tickers = pd.read_csv(
            filepath_or_buffer='Data/TickersFIIs.csv', encoding='iso-8859-1', sep=';')
        return tickers

    def colect_data_from_ticker(self, tickers) -> pd.DataFrame:
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
                    f"[[blue]{datetime.now().strftime('%H:%M:%S')}[/]] ---> [[yellow]Dados coletados para o ativo[/]] :: [{row['Ticker']}] [[yellow]Empresas[/]] :: [{qtd_index} de {len(tickers)}]")

            except selenium.common.exceptions.NoSuchElementException as e:
                console.print(
                    f"[[red]Erro[/]] Não foi possível coletar dados do ativo {row['Ticker']} :: {str(e)}")

            except KeyboardInterrupt:
                break

            qtd_index += 1

        return df

    @staticmethod
    def tratar_dados(df: pd.DataFrame) -> pd.DataFrame:
        df = df.applymap(lambda x: x.replace('%', ''))
        df = df.applymap(lambda x: x.replace('.', '').replace(',', '.'))
        df = df.applymap(lambda x: 0 if x == "-" else x)

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
            (df['DY (12 Meses)'] >= config['DY Min']) & \
            (df['P/VP'] > config['PVP Min']) & \
            (df['P/VP'] <= config['PVP Max']) & \
            (df['N° de cotistas'] >= config['N de cotistas Min']) & \
            (df['Liquidez média diária'] >= config['Liquidez média diária Min']) & \
            (df['Valorização (12 Meses)'] > config['Valorização (12 Meses) Min'])

        df = df[param_filtro_]

        return df

    @staticmethod
    def ahp_gaussian(df: pd.DataFrame) -> pd.DataFrame:
        columns_verif = [
            'DY (12 Meses)', 'P/VP', 'N° de cotistas', 'DY CAGR (3 Anos)',
            'Liquidez média diária', 'Rendimento mensal médio (24 Meses)',
            'Valorização (12 Meses)', 'Valorização (Mês atual)', 'Valor em Caixa'
        ]

        # Definição dos critérios de maximização e minimização
        maximize_criteria = [
            'DY (12 Meses)', 'DY CAGR (3 Anos)', 'Liquidez média diária', 'N° de cotistas',
            'Valorização (12 Meses)', 'Valorização (Mês atual)', 'Rendimento mensal médio (24 Meses)',
        ]

        minimize_criteria = [
            'P/VP', 'Valor em Caixa'
        ]

        # Criação da matriz normalizada método AHP-GAUSSIANO
        df_normalized = df[columns_verif].copy()

        # Para critérios de maximização, utiliza-se a fórmula de normalização
        for col in maximize_criteria:
            df_normalized[col] = (df[col] - df[col].min()) / \
                (df[col].max() - df[col].min())

        # Para critérios de minimização, utiliza-se a fórmula de normalização inversa
        for col in minimize_criteria:
            df_normalized[col] = 1 / \
                ((df[col] - df[col].min()) / (df[col].max() - df[col].min()))

        # Calcular a média, desvio padrão e o fator gaussiano
        df_normalized['Mean'] = df_normalized.mean(axis=1)
        df_normalized['StdDev'] = df_normalized.std(axis=1)
        df_normalized['Gaussian_Factor'] = df_normalized['StdDev'] / \
            df_normalized['Mean']

        # Análise de sensibilidade do ranking do método AHP-GAUSSIANO
        df_normalized['Ranking_Sensitivity'] = df_normalized['Gaussian_Factor'].sum(
        ) - df_normalized['Gaussian_Factor']

        # Ranking Final
        df_normalized['Ranking_Final'] = df_normalized['Ranking_Sensitivity'].rank(
            ascending=False)

        # Ordenando o dataframe pelo ranking final
        df_normalized.sort_values(by='Ranking_Final', inplace=True)

        for index,  row in df_normalized.iterrows():
            df.loc[index, 'Ranking AHP Gaussiano'] = row['Ranking_Final']

        df.sort_values(by='Ranking AHP Gaussiano',
                       ascending=True, inplace=True)

        return df

    @staticmethod
    def definir_quantidade_de_compra(df: pd.DataFrame, ammount: int = 1000) -> pd.DataFrame:
        df = df
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

        return df


if __name__ == "__main__":
    app = FiiExplorer()
