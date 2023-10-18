import pandas as pd
import win32com.client as win32

# importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx', engine='openpyxl')

# visualizar a base de dados
pd.set_option('display.max_columns', None)

# faturamento por loja
#print(tabela_vendas[['ID Loja', 'Valor Final']]) mostando as colunas seperadamente
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum() #Agrupando as lojas
print(faturamento)

# quantidade de produtos vendidos por loja
qtdade_produtos = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(qtdade_produtos)

print('-' *50)
# ticket médio por porduto em cada loja
ticket_medio = (faturamento['Valor Final'] / qtdade_produtos['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

# enviar um email com o relatório (import pywin32)

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To ='seuemail@gmail.com'
mail.Subject = 'Relatório de vendas por loja'
mail.HTMLBody = f'''
<h2>Relatório de vendas</h2>

<p>Prezado,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</P>
{qtdade_produtos.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Att.,</p>
<p>Richard</p>
'''

mail.Send()
print('Enviado')