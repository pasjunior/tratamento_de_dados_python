# Tratamento de Dados
[![NPM](https://img.shields.io/npm/l/react)](https://github.com/pasjunior/tratamento_de_dados_python/blob/main/licence)

# Descrição do projeto
Este código é um script Python que lê arquivos de Excel contendo informações de processos judiciais e realiza tratamento dessas informações.

## Bibliotecas utilizadas
* pandas: biblioteca para manipulação de dados em formato de tabela;
* re: biblioteca para expressões regulares;
* tkinter: biblioteca para construir interfaces gráficas em Python;
* os: biblioteca para trabalhar com sistema operacional;
* locale: biblioteca para formatação de números.

## Funcionamento do código

O script começa importando as bibliotecas necessárias e criando uma janela vazia do tkinter. Em seguida, é aberta uma caixa de diálogo para o usuário selecionar os arquivos de Excel contendo as informações dos processos judiciais a serem tratados.

Após selecionar os arquivos, o código utiliza a biblioteca pandas para ler os arquivos e manipular as informações. Para cada arquivo selecionado, é lido a planilha "BASE" e selecionadas algumas colunas .

Em seguida, são realizadas algumas operações de formatação de dados, tais como:

conversão de dados para string;
extração de informações específicas de algumas colunas utilizando expressões regulares;
formatação de números no padrão brasileiro utilizando a biblioteca locale;
substituição de valores nulos por zeros em algumas colunas;
substituição do caractere "-" por "" em algumas colunas;
conversão do formato da data na coluna "DATA DISTRIBUIÇÃO" de "yyyy-mm-dd" para "dd/mm/yy";
renomeação das colunas.
Por fim, o código salva as informações tratadas em um novo arquivo Excel com o nome "_output" adicionado ao final do nome original do arquivo.

# Como utilizar
Para utilizar este script, é necessário ter o Python e as bibliotecas pandas, tkinter, os e locale instaladas no computador.

Após clonar ou baixar este repositório, basta executar o arquivo tratar_processos_judiciais.py e selecionar os arquivos de Excel contendo as informações dos processos a serem tratados na caixa de diálogo que será aberta.

Ao final do processamento, serão gerados novos arquivos de Excel com as informações tratadas na mesma pasta dos arquivos originais.

# Contribuições
Contribuições para o projeto são sempre bem-vindas! Caso você queira sugerir melhorias, correções de bugs ou novas funcionalidades, por favor, abra uma issue ou pull request.
