import win32com.client as win32
from win32com.client import Dispatch

def recuperar_emails(remetente):
    try:
        outlook = Dispatch('outlook.application').GetNamespace('MAPI')
        inbox = outlook.GetDefaultFolder(3)
        caixa_entrada = outlook.GetDefaultFolder(6)
        emails = inbox.items

        emails_recuperados = 0
        emails_verificados = 0

        print(f"A quantidade de emails excluidos é {len(emails)}")
        print(f"Começando a recuperar emails do remetente {remetente}")
        
        for email in emails:
            try:
                emails_verificados += 1
                remetente = email.SenderEmailAddress
                if remetente in remetentes_a_recuperar:
                    print(f"Recuperando email {email}")
                    email.Move(caixa_entrada)
                    emails_recuperados += 1
                    print(f"O email {email} foi recuperado")
    
            except Exception as e:
                print(f"Ocorreu um problema, {e}")
    except Exception as e:
        print(f"Ocorreu um problema, {e}")
    print(f"Processo Concluido.")

    emails_vari_recuperdos = 'email' if emails_recuperados == 1 else 'emails'
    emails_vari_verificados = 'email' if emails_verificados == 1 else 'emails'

    print(f"Foram recuperados {emails_recuperados} {emails_vari_recuperdos}, e foram verificados {emails_verificados} {emails_vari_verificados}.")

#Preencher a lista com os Remetentes que os E-mails serão recuperados.
remetentes_a_recuperar = []

if __name__ == '__main__':  
    recuperar_emails(remetentes_a_recuperar)