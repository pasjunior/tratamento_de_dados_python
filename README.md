# Tratamento de Dados
[![NPM](https://img.shields.io/npm/l/react)](https://github.com/pasjunior/tratamento_de_dados_python/blob/main/licence)

# Descrição do projeto
Esse script realiza o tratamento de bases de dados de pedidos e fechamentos, que são fornecidas em formato Excel, com o objetivo de deixá-las em um formato padronizado e pronto para análises.

## Bibliotecas utilizadas
* pandas: biblioteca para manipulação e análise de dados
* re: biblioteca para expressões regulares
* tkinter: biblioteca gráfica para construção de interfaces gráficas
* time: biblioteca para trabalhar com tempo
* os: biblioteca para interface com o sistema operacional
* locale: biblioteca para formatação de números no padrão brasileiro

## Funcionamento do código

1. O usuário é solicitado a selecionar os arquivos de pedidos e fechamentos por meio de caixas de diálogo.

2. Os dados do arquivo de pedidos são lidos e armazenados em um objeto DataFrame do pandas. Em seguida, são feitas as seguintes operações:
 * A última linha do arquivo é removida, pois geralmente contém informações adicionais que não são relevantes.
 * A coluna 'Objeto' é limpa de espaços em branco.
 * É extraída uma chave única para cada processo a partir da coluna 'Pasta_Websijur'.
 * A coluna 'Criado_Em' é formatada para o padrão brasileiro.
 * O arquivo é exportado para um novo arquivo Excel.
3. Os dados dos arquivos de fechamento são lidos e armazenados em um objeto DataFrame do pandas. Em seguida, são feitas as seguintes operações:
 * Todas as colunas são convertidas para o tipo string para evitar erros de formatação.
 * É extraída uma chave única para cada processo a partir da coluna 'PASTA'.
 * Os valores das colunas 'VALOR ATUALIZAÇÃO PROVÁVEL', 'PEDIDO ATUALIZADO', 'PROVÁVEL ATUALIZADO' e 'PROVÁVEL MÊS ANTERIOR' são formatados no padrão brasileiro.
 * Os valores faltantes nas colunas mencionadas acima são preenchidos com zero.
 * É realizada uma substituição do caractere '-' por zero nas colunas mencionadas acima.
 * A coluna 'DATA DISTRIBUIÇÃO' é convertida para o tipo datetime.
4. O resultado final é um objeto DataFrame do pandas para cada base de dados, que estão prontos para serem utilizados em análises.

# Como utilizar
Para utilizar este script, é necessário ter o Python e as bibliotecas pandas, tkinter, os e locale instaladas no computador.

Após clonar ou baixar este repositório, basta executar o arquivo tratar_processos_judiciais.py e selecionar os arquivos de Excel contendo as informações dos processos a serem tratados na caixa de diálogo que será aberta.

Ao final do processamento, serão gerados novos arquivos de Excel com as informações tratadas na mesma pasta dos arquivos originais.

# Contribuições
Contribuições para o projeto são sempre bem-vindas! Caso você queira sugerir melhorias, correções de bugs ou novas funcionalidades, por favor, abra uma issue ou pull request.
