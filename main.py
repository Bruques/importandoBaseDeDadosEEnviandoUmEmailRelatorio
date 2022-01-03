import pandas as pd
#import win32com.client as win32

# Importar a base de dados (arquivo excel)
tabelaVendas = pd.read_excel('Vendas.xlsx')

# Visualizar a base de dados
# O codigo abaixo pede para o pycharm mostrar todas as colunas da tabela
pd.set_option('display.max_columns', None)

# Faturamento por loja
# Entre os dois [] filtramos qual colunas nos queremos
# e o .groupby() nos selecionamos oq queremos agrupar, que no caso e o ID Loja
# o .sum() pedimos para que seja somada o Valor Final
faturamento = tabelaVendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()

# Quantidade de produtos vendidos por loja

quantidade = tabelaVendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()

# Ticket medio por produto em cada loja
# O ticket medio e o valor do faturamento / quantidade
# Entao selecionamos a variavel e entre [] colocamos o que nos queremos filtrar das colunas
# Mas o python nao vai nos entregar uma tabela, entao nos tems que pedir pra que ele faca isso com o ".to_frame()"
ticketMedio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticketMedio = ticketMedio.rename(columns={0: 'Ticket Médio'})
print(ticketMedio)



# Enviar um email com o relatorio

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'colocarAquiUmEmail@gmail.com'
mail.Subject = 'Ex.: Relatório de vendas por loja'
mail.HTMLBody = f'''
<p>Prezados</p>

<p>Faturamento</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade de vendas</p>
{quantidade.to_html()}

<p>Ticket médio dos produtos em cada loja:</p>
{ticketMedio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer duúvida estou a disposição</p>

<p>Att.,</p>
<p>Bruno Nascimento Marques</p>

'''

mail.Send()
print('Email enviado')