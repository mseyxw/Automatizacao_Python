import os
import win32com.client

def salvar_anexos(email, pasta_destino):
    attachments = email.Attachments
    for attachment in attachments:
        file_path = os.path.join(pasta_destino, attachment.FileName)
        attachment.SaveAsFile(file_path)

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

pasta_origem = outlook.GetDefaultFolder(6).Items

endereco_remetente = "mariasimbrl@gmail.com"

pasta_destino = r'C:\Users\maria\OneDrive\Documentos\Anexos salvos'

for email in pasta_origem:
    if email.SenderEmailAddress.lower() == endereco_remetente.lower():
        if email.Attachments.Count > 0:
            salvar_anexos(email, pasta_destino)

print("Anexos salvos com sucesso!")





