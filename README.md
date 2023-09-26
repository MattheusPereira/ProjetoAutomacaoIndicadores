# Projeto Automação de indicadores

Foi utilizado as bibliotecas do Pandas, Pathlib e Win32com.client para realização do projeto.
 
# Objetivo: 
Treinar e criar um Projeto Completo que envolva a automatização de um processo feito no computador

# Descrição:
Na pasta Base de Dados, possui o arquivo vendas.xlsx onde nele consta as vendas de 25 lojas espalhadas pelo Brasil, com isso é feito o tratamento dos dados para realização de um resumo simples e direto ao ponto, usado pela equipe de gerência de loja para saber os principais indicadores de cada loja e permitir em 1 página (daí o nome OnePage) tanto a comparação entre diferentes lojas, quanto quais indicadores aquela loja conseguiu cumprir naquele dia ou não.

# Arquivos e Informações Importantes
- Arquivo Emails.xlsx com o nome, a loja e o e-mail de cada gerente. Obs: Sugiro substituir a coluna de e-mail de cada gerente por um e-mail seu, para você poder testar o resultado
- Arquivo Vendas.xlsx com as vendas de todas as lojas. Obs: Cada gerente só deve receber o OnePage e um arquivo em excel em anexo com as vendas da sua loja. As informações de outra loja não devem ser enviados ao gerente que não é daquela loja.
- Arquivo Lojas.csv com o nome de cada Loja
- Ao final, sua rotina deve enviar ainda um e-mail para a diretoria (Neste caso será necessário estar instalado o Microsoft Office Outlook e estar configurado com seu email para poder enviar os emails) com 2 rankings das lojas em anexo, 1 ranking do dia e outro ranking anual. Além disso, no corpo do e-mail, deve ressaltar qual foi a melhor e a pior loja do dia e também a melhor e pior loja do ano. O ranking de uma loja é dado pelo faturamento da loja.
- As planilhas de cada loja devem ser salvas dentro da pasta da loja com a data da planilha, a fim de criar um histórico de backup

# Indicadores do OnePage
- Faturamento -> Meta Ano: 1.650.000 / Meta Dia: 1000
- Diversidade de Produtos (quantos produtos diferentes foram vendidos naquele período) -> Meta Ano: 120 / Meta Dia: 4
- Ticket Médio por Venda -> Meta Ano: 500 / Meta Dia: 500

Obs: Cada indicador deve ser calculado no dia e no ano. O indicador do dia deve ser o do último dia disponível na planilha de Vendas (a data mais recente)

Obs2: Dica para o caracter do sinal verde e vermelho: pegue o caracter desse site (https://fsymbols.com/keyboard/windows/alt-codes/list/) e formate com html

​
