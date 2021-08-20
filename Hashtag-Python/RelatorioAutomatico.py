from email.message import MIMEPart
import pandas as pd
import smtplib
import ssl
from email.mime.text import MIMEText

# importar a base de dados
planilha = 'Vendas.xlsx'
tabela_vendas = pd.read_excel(planilha)

# Visualizar a base de dados
pd.set_option('display.max_columns', None)
print('Tabela Completa:')
print(tabela_vendas)

# Calcular Faturamento por loja
faturamento = tabela_vendas[[
    'ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print('Tabela de Faturamento')
print(faturamento)

# Quantidade de produtos vendidos por loja
produtos_por_loja = tabela_vendas[[
    'ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print('Produtos por Loja: ')
print(produtos_por_loja)

# Ticket Médio por produto em cada loja (Faturamento / Qtde Vendida)
# Transforma objeto da divisão em tabela.
print('-='*50)
print('Ticket Medio')
ticket_medio = (faturamento['Valor Final'] /
                produtos_por_loja['Quantidade']).to_frame()

print(ticket_medio)
# Configurar e-mail
login = input('Digite seu login do google: ')
senha = input('Digite sua senha do Google: ')

sender = 'leoskinner@gmail.com'
receivers = ['leo_skinner@hotmail.com', 'leksinfo@gmail.com']
body_of_email = f'''<p>Bom dia!</p>

<p>Segue relatório automatizado pelo python.</p>

<p>Faturamento:</p>
{faturamento.to_html()}

<p>Produtos por loja:</p>
{produtos_por_loja.to_html()}

<p>Ticket médio dos produtos por loja:</p>
{ticket_medio.to_html()}

<p>Att,</p>
<p>Leonardo Skinner</p>

'''

msg = MIMEText(body_of_email, 'html')
msg['Subject'] = 'Teste Skinner: Python enviando e-mail'
msg['From'] = sender
msg['To'] = ','.join(receivers)
# Enviar e-mail com relatório
try:
    s = smtplib.SMTP_SSL(host='smtp.gmail.com', port=465)
    s.login(user=login, password=senha)
    s.sendmail(sender, receivers, msg.as_string())
    s.quit()
    print('email enviado!')
except:
    print('Something went wrong....')
