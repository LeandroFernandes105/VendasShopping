import pandas as pd
import win32com.client as win32



#Improtar a base de dados

tabela_vendas = pd.read_excel('Vendas.xlsx')




#Visulizar a base de dados

pd.set_option('display.max_columns',None) # serve para mostrar todas as colunas
print(tabela_vendas)

print ('-' *50)
#Faturamento por loja
faturamento=tabela_vendas[['ID Loja','Valor Final']].groupby('ID Loja').sum() #filtra colunas IdLoja valor final e agrupa por Id loja somando os valores de Valor Final

print(faturamento)
#Quantidade de produtos vendidos por loja

quantidade=tabela_vendas[['ID Loja','Quantidade']].groupby('ID Loja').sum() #filtra colunas IdLoja valor final e agrupa por Id loja somando os valores de Valor Final

print(quantidade)

#ticket médio por produto em cada loja.
#Faturamento da loja / quantiade de produto vendido pela loja 

print ('-' *50)
# ticket= faturamento/quantidade  não posso fazer assim
ticket_medio =(faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio=ticket_medio.rename(columns={0: 'Ticket Médio'}) #mudou o nome da coluna 0 pata Ticket médio

print(ticket_medio)

#Enviar um email com o relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'lfernandes1898@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas de cada loja.</p>

<p><b>Faturamento:</b></p>
{faturamento.to_html(formatters={'Valor Final':'R${:,.2f}'.format})} 

<p><b>Quantidade Vendida:</b></p>
{quantidade.to_html()}

<p><b>Ticket médio dos Produtos em cada loja:</b></p>
{ticket_medio.to_html(formatters={'Ticket Médio':'R${:,.2f}'.format})}

<p>Qualquer dúvida entre em contato.</p>

<p>Att..</p>
<p>Leandro</p>
'''

#formata a tabela Valor final  : toda formatação 
#começa com : , para seprar milhar . para separar decimal 2f é para 2 caas decimal 
mail.Send()
print('Email enviado')





#openpyxl pacote para ler excel  py -m pip install openpyxl

#pywmin32 para mandar email