import win32com.client as win32
import pandas as pd
from tkinter import messagebox


# importar base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

# visualizar a base de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)

# faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

# quantidade de produto vendido por loja
prod_vendido = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(prod_vendido)

print('-' * 50)
# ticket medio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / prod_vendido['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

def message():
    messagebox.showinfo(title="info", message="Email enviado!")

# enviar email com o relatorio
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'nathmorais223@gmail.com' # email de destino
mail.Subject = 'Relatorio de Vendas por Loja'
mail.HTMLBody = f''' 
<p>Nathalia Cristinne,</p>

<p>Segue o Relatório de vendas por cada loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{prod_vendido.to_html(formatters={'Quantidade': '{:,}'.format})}

<p>Ticket Médio dos produtos em cada loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}


<p>Qualquer dúvida estou à disposição.</p>

<p>Att.,</p>
<p>Mateus Torquato</p>
'''


mail.Send()


message()