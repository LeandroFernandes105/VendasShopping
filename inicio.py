import pandas as pd
import win32com.client as win32



#Importar a base de dados

tabela_vendas = pd.read_excel('Vendas.xlsx')


#Visulizar a base de dados

pd.set_option('display.max_columns',None) 
print(tabela_vendas)

print ('-' *50)
#Faturamento por loja
faturamento=tabela_vendas[['ID Loja','Valor Final']].groupby('ID Loja').sum().reset_index() 

print(faturamento)
#Quantidade de produtos vendidos por loja

quantidade=tabela_vendas[['ID Loja','Quantidade']].groupby('ID Loja').sum().reset_index() 

print(quantidade)

#ticket médio por produto em cada loja.
#Faturamento da loja / quantidade de produto vendido pela loja 

print ('-' *50)

ticket_medio =(faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio=ticket_medio.rename(columns={0: 'Ticket Médio'}) 
# Unindo o faturamento e a quantidade para incluir ID na tabela do ticket médio
ticket_medio = faturamento[['ID Loja']].merge(ticket_medio, left_index=True, right_index=True)

print(ticket_medio) 


#Enviar um email com o relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'email@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas de cada loja.</p>

<p><b>Faturamento:</b></p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format}, index=False)} 

<p><b>Quantidade Vendida:</b></p>
{quantidade.to_html(index=False)}

<p><b>Ticket médio dos Produtos em cada loja:</b></p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format}, index=False)}

<p>Qualquer dúvida entre em contato.</p>

<p>Att..</p>
<p>Leandro</p>
'''


mail.Send()
print('Email enviado')







