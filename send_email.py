import pandas as pd
# from appscript import app, k # caso queira enviar email com MacOs
# import win32com.client as win32 # caso queira enviar email com windows

# ler dados da planilha
planilha = pd.read_excel("Vendas.xlsx")
pd.set_option('display.max_columns', None)

# filtrar dados da planilha
faturamento = planilha[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
quantidade = planilha[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Medio'})

### enviar no email outlook para MacOs

# Conecta ao aplicativo Microsoft Outlook do Mac
# outlook = app("Microsoft Outlook")
# email = outlook.make(
#     new=k.outgoing_message,
#     with_properties={
#         k.subject: "Relatório de Vendas completo",
#         k.content: f"""
#         <h2>Prezados,</h2>
#         <p>Segue o relatório de vendas completo.</p>
#         <p>Faturamento</p>
#         {faturamento.to_html(formatters={"Valor Final": 'R${:,.2f}'.format})}
#         <p>Quantidade</p>
#         {quantidade.to_html()}
#         <p>Ticket medio</p>
#         {ticket_medio.to_html(formatters={"Ticket Medio": 'R${:,.2f}'.format})}
#         <p>Cordialmente,</p>
#         <p>Felipe Duarte</p>
#         """
#     }
# )
# email.to_recipients.end.make(
#     new=k.recipient,
#     with_properties={
#         k.email_address: {
#             k.address: "duarte.dev@outlook.com"
#         }
#     }
# )
# email.send()
# print("E-mail enviado via Outlook no Mac!")


### enviar no email outlook para Windows

# # Criar a conexão com o Outlook
# outlook = win32.Dispatch("outlook.application")
# # Criar um novo e-mail (0 representa um item de e-mail)
# mail = outlook.CreateItem(0)
# # Configurar os campos do e-mail
# mail.To = "destinatario@exemplo.com.br"
# mail.Subject = "Assunto do e-mail automático"
# mail.Body = "Olá, este é um e-mail enviado automaticamente usando Python e Outlook."
# # Opcional: Corpo em formato HTML
# mail.HTMLBody = """
# <p>Olá,</p>
# <p>Este é um e-mail com <b>formatação HTML</b> enviado via Python.</p>
# """
# # Opcional: Adicionar um anexo
# # anexo = r"C:\caminho\para\o\seu\arquivo.pdf"
# # mail.Attachments.Add(anexo)
# # Enviar o e-mail
# mail.Send()
# print("E-mail enviado via Outlook no Windows!")
