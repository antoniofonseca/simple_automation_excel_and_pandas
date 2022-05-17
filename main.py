import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# 1 - importar base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

# 2 - visualizar base de dados
pd.set_option('display.max_columns', None)
# print(tabela_vendas[['ID Loja', 'Valor Final']])

# 3- faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
# print(faturamento)

# 4- quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
# print(quantidade)

# 5- 'ticket' médio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
# print(ticket_medio)

# 6- enviar um correio eletrônico com o relatório

# Server parameters
smtp_ssl_host = 'smtp.gmail.com'
smtp_ssl_port = 465
username = 'name@email.com'
password = 'password_token'

# Sender and receiver
me = "sender@email.com"
you = "receiver@email.com"

# Create message container - the correct MIME type is multipart/alternative.
msg = MIMEMultipart('alternative')
msg['Subject'] = "Link"
msg['From'] = me
msg['To'] = you

# Create the body of the message (a plain-text and an HTML version).
# text = "Hi!\nHow are you?\nHere is the link you wanted:\nhttp://www.python.org"
html = f'''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html(formatters={'Quantidade': '{:,}'.format})}

<p>Ticket Médio dos Produtos em casa Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>

<p>Att.,</p>
<p>ASF</p>
'''

# Record the MIME types of both parts - text/plain and text/html.
# part1 = MIMEText(text, 'plain')
part2 = MIMEText(html, 'html')

# Attach parts into message container.
# According to RFC 2046, the last part of a multipart message, in this case
# the HTML message, is best and preferred.
# msg.attach(part1)
msg.attach(part2)
# Send the message via local SMTP server.
mail = smtplib.SMTP_SSL(smtp_ssl_host, smtp_ssl_port)
mail.ehlo()
mail.login(username, password)
mail.sendmail(me, you, msg.as_string())
mail.quit()

print('E-mail enviado com sucesso!')
