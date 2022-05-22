#!/usr/bin/env python
# coding: utf-8

# In[ ]:


from imap_tools import MailBox, AND

# pegar emails de um remetente para um destinatário
username = "seu_email"
password = "senha"

# lista de imaps: https://www.systoolsgroup.com/imap/
meu_email = MailBox('imap.gmail.com').login(username, password)

# criterios: https://github.com/ikvk/imap_tools#search-criteria
lista_emails = meu_email.fetch(AND(from_="remetente", to="destinatario")) 
for email in lista_emails:
    print(email.subject)
    print(email.text)


# pegar emails com um anexo específico
lista_emails = meu_email.fetch(AND(from_="remetente"))
for email in lista_emails:
    if len(email.attachments) > 0:
        for anexo in email.attachments:
            if "TituloAnexo" in anexo.filename:
                print(anexo.content_type)
                print(anexo.payload)
                with open("Teste.xlsx", 'wb') as arquivo_excel:
                    arquivo_excel.write(anexo.payload)

