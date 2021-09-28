
import pandas as pd
import subprocess

import smtplib
from email.mime.text import MIMEText
from email.header import Header

# Importar base De Dados:
tabela_vendas = pd.read_excel('Vendas.xlsx')
""" print(tabela_vendas)"""

# Vizualizar Base De Dados:
# -> (Opção, Valor)  _>Valor none == Não tem máximo de coluans.
# .. Não oculta mais coluna nenuma
pd.set_option('display.max_columns', None)
# Uma Coluna == Um []
# print(tabela_vendas['Produto'])

# Duas DUAS Colunas OU Mais == Dois [[]]
# print(tabela_vendas[['Produto', 'Valor Final', 'Quantidade']])


# Faturamento Por Loja:
"""Com este método -> Ele Não Repete os Nomes em 'ID Loja' Então, toda vez
 que o Nome de Uma mesma Loja aparece em 'Id Loja' o Valor que esta referenciado
  na coluna (Neste Caso) 'Valor Final' É somado Numa Mesma Linha de 'Quantidade'
   Antes:
   ID Loja| Quantidade
   Loja_X | 1
   Loja_X | 1
   Loja_Y | 4
   Loja_Y | 0
   
   Depois:
   Loja_X | 2
   Loja_Y | 4
   
   Não se repetem mais."""
# Selecionou duas colunas ID & Quant  .groupby('Linha Que Não se Repete Mais').somar Valores da Coluna a frente (Valor Final)
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()  # OUmean()
#print(faturamento)


# Quantidade de produtos vendidos por cada Uma das Loja:
#print(tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum())
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()



# Ticket Médio:
# -> Faturamento Dividio Pela Quantidade De Produtos Vendidos por Loja

#SÓ FUNCIONOU COM 1 COUCHETE SÓ []   ASSIM [[]]  NÃO (!)
"""Sempre que fazer uma operação matematica entre tabelas. É preciso Convertelas De volta P/ Tabela Usando -> Frame"""
#.. e Cercar a Linha com '()'
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()

#Nome_da_Tabela Recebe Ela Mesma.Renomeia(Colna=Nome_AntigoDa_Colna: 'Nome Novo'})
ticket_medio = ticket_medio.rename(columns={0: 'Ticket_Médio'})

#print(ticket_medio)





# ENVIAR E-MAIL COM RELATÓRIOS:
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

html = f"""\
<p> Prezados, </p>

<p> Faturamento: </p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p> Quantidade Vendida </p>
{quantidade.to_html()}

<p> Ticket Medio Dos Produtos em Cada Loja </p>
{ticket_medio.to_html(formatters={'Ticket_Médio': 'R${:.2f}'.format})}

<p> Qualquer Duvida estou a disposição. </p>

<p> Att.,</p>
"""

# conexão com os servidores do google
smtp_ssl_host = 'smtp.gmail.com'
smtp_ssl_port = 465

# username ou email para logar no servidor GOOGLE:
senha = str(input("Senha do E-mail: "))
username = 'diksonnn@gmail.com'
password = senha







from_addr = 'diksonnn@gmail.com'
to_addrs = ['dikson_santos@outlook.com', 'dikson.enterprise@gmail.com']

# a biblioteca email possuí vários templates
# para diferentes formatos de mensagem
# neste caso usaremos MIMEText para enviar
# somente texto (TIVE QUE CONVERTER P/ STR )
Final = ticket_medio,  quantidade,  faturamento

#message = MIMEText(str(Final))
#message = MIMEText(str(Final))
message = MIMEText(html, 'html')
message['subject'] = 'Relatorios (3))'
message['from'] = from_addr
message['to'] = ', '.join(to_addrs)

# conectaremos de forma segura usando SSL
server = smtplib.SMTP_SSL(smtp_ssl_host, smtp_ssl_port)
# para interagir com um servidor externo precisaremos
# fazer login nele
server.login(username, password)
server.sendmail(from_addr, to_addrs, message.as_string())

#message.attach(part1)
#message.attach(part2)

server.quit()
print('E-Mail Enviado')


"""

Use este Link Para desativar a segurança (Da conta de G-Mails) contra aplicativos não 
tão seguros:

https://myaccount.google.com/lesssecureapps?pli=1&rapt=AEjHL4OFPGQqQGrZXltEMUmUXNkldDLGMe56C8M2YocIigtbr5k75o-cExl75tBAhCDydu_lLplhdvFKcjuaTWBcBN_ON448uQ
"""

