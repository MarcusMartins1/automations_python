# Importa a biblioteca pandas para manipulação de dados
import pandas as pd
# Importa classes necessárias para compor e enviar e-mails
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib

# Carrega o arquivo Excel contendo os dados das pessoas
pessoas = pd.read_excel("pessoas.xlsx")

# Loop através de cada linha do DataFrame 'pessoas'
for index, pessoas in pessoas.iterrows():
    # Cria um objeto MIMEMultipart para compor o e-mail
    msg = MIMEMultipart()
    # Define o assunto do e-mail
    msg['Subject'] = 'Email teste'
    # Define o remetente do e-mail
    msg['From'] = 'seuemailjr@gmail.com'
    # Define o destinatário do e-mail
    msg['To'] = pessoas['email']
    
    # Cria a mensagem a ser enviada no corpo do e-mail
    message = f"Opa, passando pra avisar que a automação de {pessoas['nome ']} foi {pessoas['Assunto']}"
    # Anexa a mensagem ao objeto MIMEMultipart
    msg.attach(MIMEText(message, 'plain'))

    # Configura o servidor SMTP para enviar o e-mail
    server = smtplib.SMTP('smtp.gmail.com', port=587)
    # Inicia a comunicação TLS (Transport Layer Security) com o servidor SMTP
    server.starttls()
    # Faz login no servidor SMTP com o e-mail e senha fornecidos
    server.login('seuemail@gmail.com', 'senhadanet')
    # Envia o e-mail do remetente para o destinatário
    server.sendmail(msg['From'], msg['To'], msg.as_string())
    # Encerra a conexão com o servidor SMTP
    server.quit()
