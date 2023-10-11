# _The Aurorion_ <a href="https://www.python.org/downloads/release/python-3120/"><img title="Version" src="https://img.shields.io/badge/Python-3.12.0-yellow.svg?style=flat-square"></a> 

## Introdução ao programa

A palavra 'Aurorion': Inspiração para Riqueza e Excelência

'Aurorion' incorpora o prefixo 'Aur', que, em latim, está associado ao ouro, ao brilho e ao nascer do sol. Essa conexão com a riqueza e a luz adiciona uma dimensão de valor e sucesso ao nome. O sufixo 'Orion' nos faz lembrar da majestosa constelação que tem sido símbolo de grandeza e orientação por séculos. 'Aurorion' convida a explorar um caminho de riqueza e excelência, fornecendo uma bússola para a direção certa.

Ao embarcarmos na jornada de criar e dar vida a um projeto, 'Aurorion' surge como uma escolha inspiradora. Com um nome que evoca a visão de riqueza, luz e orientação, 'Aurorion' representa não apenas um nome, mas um símbolo de promessa e potencial ilimitado.

Este script coleta e analisa dados de ativos financeiros, mais especificamente fundos de investimento imobiliário (FIIs) no Brasil. Ele extrai informações do site 'statusinvest.com.br' utilizando a biblioteca Selenium e, em seguida, processa e filtra os dados para análises posteriores. Aqui está uma visão geral do código e de suas funcionalidades:

## Objetivo:

O principal objetivo deste script é:

1. Coletar dados para uma lista de ativos financeiros (FIIs) no site 'statusinvest.com.br'.
2. Calcular vários indicadores financeiros para cada ativo.
3. Aplicar um conjunto de filtros e regras predefinidas para fazer recomendações de investimento com base nos dados coletados.
4. Classificar os ativos com base em critérios específicos.
5. Determinar a quantidade a ser comprada para cada ativo com base em uma estratégia de alocação predefinida.
6. Verificar se algum ativo deve ser adicionado ou removido de uma carteira de investimentos existente.

## Dependências Principais:

O script depende de várias bibliotecas e módulos Python, incluindo:

- `os`: Funções do sistema operacional.
- `json`: Para ler arquivos de configuração.
- `time`: Para introduzir atrasos.
- `numpy` e `pandas`: Para manipulação e análise de dados.
- `yfinance` (yf): Para obter dados de histórico de preços de ações.
- `selenium`: Para automatizar a coleta de dados na web.
- `ChromeDriverManager`: Para gerenciar o Chrome WebDriver.
- `rich`: Para formatação de saída na console com texto formatado.

## Estrutura do Script:

O script está estruturado em uma classe Python, `FundExplorer`, que contém diversos métodos para tarefas diferentes:

1. `coletar_tickers`: Coleta uma lista de ativos financeiros a partir de um arquivo CSV que contém tickers e outras informações.
2. `coletar_dados_do_ticker`: Faz a coleta de dados no 'statusinvest.com.br' para cada ativo na lista, incluindo diversas métricas financeiras, como preço, dividendos e mais.
3. `tratar_dados`: Processa e limpa os dados coletados, convertendo strings em tipos de dados adequados e arredondando valores numéricos.
4. `indicadores_tecnicos`: Calcula indicadores técnicos, como volatilidade, RSI (Índice de Força Relativa) e retornos simples para cada ativo.
5. `filtro`: Aplica critérios de filtragem predefinidos para gerar recomendações de investimento (Comprar, Manter ou Não Comprar).
6. `ranking`: Classifica os ativos com base em critérios específicos e calcula uma pontuação total para cada ativo.
7. `definir_quantidade_de_compra`: Determina a quantidade a ser comprada para cada ativo com base em uma estratégia de alocação predefinida.
8. `verificar_investimentos`: Verifica se algum ativo precisa ser adicionado ou removido de uma carteira de investimentos existente.

A função `main` orquestra todo o fluxo de trabalho, criando uma instância de `FundExplorer` e chamando esses métodos em sequência.

## Uso:

Para usar este script, siga estas etapas:

1. Certifique-se de ter as bibliotecas Python necessárias instaladas. Você pode usar o arquivo `requirements.txt` para instalá-las.
2. Defina seus critérios de investimento no arquivo `config/FundConfig.json`.
3. Execute o script executando-o em seu ambiente Python.

## Saída:

O script gera uma lista filtrada de ativos financeiros, juntamente com recomendações e classificações com base nos critérios predefinidos. Os resultados são salvos em um arquivo Excel no diretório "resultados".

Lembre-se de modificar os caminhos e configurações específicas para atender às suas necessidades antes de executar o script.
