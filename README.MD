### Requisitos
Eu tenho um arquivo do Excel uma uma lista de 2285 links com voucher do LinkedIn que se chama "vouchers.xlsx". A coluna "A" possui a nuemração sequencial, a coluna B possui os links com os vouchers, na coluna C deverá constar o status dos testes desses vouchers.

Preciso de uma automação em Python para ler a linha dessa tabela e verifique qual foi a página exibida.

Há 3 diferentes páginas que podem ser exibidas ao acessar esses links: sucesso, voucher já utilizado e número de tentativas excedida. Para cada uma mensagem dessa, eu terei um print para comparação e identificação do retorno.

Quero gravar na tabela os vouchers já utilizados e os disponíveis no coluna C, alimentando a planilha original com os resultados da busca com um simples texto: válido, utilizado. 

Quando o número de tentativas for excedida, eu preciso que o código pare de ser executado para dar continuidade posteriormente. 

Os vouchers serão escolhidos de linhas aleatórias, as quais ainda não possuem um status registrado na coluna C da tabela, dando continuidado ao preenchimento dessa consulta por um longo período, devido a impossibilidade de testar todos os códigos em uma só tentativa.

### Configuração do ambiente
pip install selenium pandas openpyxl webdriver-manager
