import win32com.client

outlook_app = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook_app.GetDefaultFolder(6)
emails = inbox.Items

def iniciar():
    palavras_chave = pegar_palavra_chaves_arquivo()

    for email in emails:
        for palavra_chave in palavras_chave:
            verificar_palavra_chave(email, palavra_chave)

def verificar_palavra_chave(email, palavra_chave):
    if palavra_chave in email.Subject:

        if not pasta_existe(palavra_chave):
            criar_pasta(palavra_chave)
        mover_email_para_pasta(email, palavra_chave)

def pasta_existe(pasta_verificar):
    for pasta in inbox.Folders:
        
        if pasta.Name == pasta_verificar:
            return True
    return False

def criar_pasta(nome_pasta):
    inbox.Folders.Add(nome_pasta)

def mover_email_para_pasta(email, palavra_chave): 
    pasta = inbox.Folders(palavra_chave)
    email.Move(pasta)

def pegar_palavra_chaves_arquivo():
    with open("palavras-chave.txt", "r", encoding='utf-8') as arquivo:
        linhas = arquivo.readlines()
        return [linha.strip() for linha in linhas]

iniciar()
