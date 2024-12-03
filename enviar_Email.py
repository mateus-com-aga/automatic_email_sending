import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os


def remover_acentos(palavra):
    # Dicionário de mapeamento de vogais com acento para vogais sem acento
    mapeamento = {
        'á': 'a', 'à': 'a', 'â': 'a', 'ã': 'a', 'ä': 'a',
        'é': 'e', 'è': 'e', 'ê': 'e', 'ë': 'e',
        'í': 'i', 'ì': 'i', 'î': 'i', 'ï': 'i',
        'ó': 'o', 'ò': 'o', 'ô': 'o', 'õ': 'o', 'ö': 'o',
        'ú': 'u', 'ù': 'u', 'û': 'u', 'ü': 'u'
    }

    # Substituir cada caractere na palavra pelo mapeamento, se existir
    palavra_sem_acentos = ''.join(mapeamento.get(letra, letra) for letra in palavra)

    return palavra_sem_acentos


def enviarEmail(email, nome):
    # Configurações do servidor SMTP
    smtp_server = 'mail.tombom.com.br'
    port = 587  # Porta para conexão com o servidor SMTP
    sender_email = ''
    password = ''

    # Configurações do e-mail
    receiver_email = email
    subject = 'Fechamento Novembro'
    body = 'Segue fechamentos em anexo.'

    # Lista de nomes de arquivo dos anexos
    file_paths = {f'convertidos/vendas_Fixa_{remover_acentos(nome.lower())}.xlsx',
                    f'convertidos/vendas_Movel_{remover_acentos(nome.lower())}.xlsx',
                    f'convertidos/vendas_Vada_{remover_acentos(nome.lower())}.xlsx'}

    # Criando o objeto Multipart
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject

    # Adicionando o corpo do e-mail
    msg.attach(MIMEText(body, 'plain'))

    # Loop sobre cada arquivo na lista de caminhos dos arquivos
    for file_path in file_paths:
        # Anexo
        filename = os.path.basename(file_path)  # Obtém apenas o nome do arquivo
        attachment = open(file_path, 'rb')

        # Criando o objeto MIMEBase
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

        # Adicionando o anexo ao e-mail
        msg.attach(part)

    # Iniciando a conexão com o servidor SMTP
    server = smtplib.SMTP(smtp_server, port)
    server.starttls()
    server.login(sender_email, password)

    # Enviando o e-mail
    text = msg.as_string()
    server.sendmail(sender_email, receiver_email, text)

    # Encerrando a conexão com o servidor SMTP
    server.quit()


nomes_unicos = ['nome_1']

emails_unicos = ['email_1']
print(nomes_unicos)

# Iterando sobre as listas de nomes e emails
for nome, email in zip(nomes_unicos, emails_unicos):
    enviarEmail(email, nome)
    print(f'Enviado com sucesso para: {nome}')
