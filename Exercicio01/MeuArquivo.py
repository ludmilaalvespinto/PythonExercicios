import pandas as pd
import win32com.client as win32

# importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

# visualizar a base de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)

# Faturamento por loja

# Demonstrar apenas as colunas que quer demonstrar
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

# Quantidade de produtos vendidos por loja
quantidadeProduto = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidadeProduto)

print('--' * 50)
# Ticket médio por produto em cada loja
ticket_medio = (faturamento ['Valor Final'] / quantidadeProduto['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

# enviar um email com o relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.to = 'ludmila.alvespinto@gmail.com'
mail.Subject = 'Relatório de vendas por loja'
mail.HTMLBody = f'''
<p>Prezados,</p> 
<br>
<p>Segue o relatório solicitado de vendas por cada loja com:</p>
<br>

<p> Faturamento:</p> 
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}


<p>Quantidade Vendida:</p>
{quantidadeProduto.to_html()}

<p>Ticket Médio dos produtos em cada loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>No mais, estou a disposição.</p>
<p>Ludmila</p>

'''

mail.Send ()
print ('Email enviado')
