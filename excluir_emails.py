import win32com.client as win32
from win32com.client import Dispatch
import traceback


def excluir_emails(remetente):
    try:
        outlook = Dispatch('outlook.application').GetNamespace('MAPI')
        inbox = outlook.GetDefaultFolder(6)
        emails = inbox.items
            
        print(f"A quantidade de emails é {len(emails)}")

        emails_excluidos = 0
        emails_verificados = 0

        for email in emails:
            try:
                emails_verificados += 1
                remetente = email.SenderEmailAddress

                if remetente in remententes_a_excluir:
                    print(f"Excluindo o email {email}, do remetente {remetente}")
                    email.Delete()
                    emails_excluidos += 1

            except Exception as e:
                print(f"Ocorreu o erro de {e}, continuando")
                continue

        print(f"\nProcessamento Concluido")
        email_variavel = 'email' if emails_excluidos == 1 else 'emails'
        print(f"Foram excluidos {emails_excluidos} {email_variavel} e foram verificados {emails_verificados} emails")

    except Exception as e:
        print("Erro ao processar outlook {e}")


remententes_a_excluir = [
'noreply@telegram.org'
]



if __name__ == "__main__":
    print("Iniciando exclusão de emails")
    excluir_emails(remententes_a_excluir)



