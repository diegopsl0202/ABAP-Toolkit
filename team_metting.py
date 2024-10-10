import os
import win32com.client
from datetime import datetime

def create_teams_meeting(subject, emails, pwo, ricef_cd, description, functional_area, signature):
    try:
        # Inicializa o Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")

        # Cria um novo e-mail
        mail_item = outlook.CreateItem(0)  # 0 corresponde a um novo e-mail

        # Define o destinatário, assunto e o corpo do e-mail
        mail_item.To = emails
        mail_item.Subject = subject

        html_body = f"""
            <html>
            <body style="font-family:Calibri, sans-serif; color: #FFFFFF; background-color: #2F2F2F;">

            <p>Hi GSS Team,</p>

            <p>Good day!</p>

            <p>Please serve this as an invite for KT Session of PWO {pwo}:</p>

                <table style="border: 1px solid #FFFFFF; border-collapse: collapse; width: 60%;">
                    <tr>
                        <th style="border: 1px solid #FFFFFF; padding: 8px;">PWO #</th>
                        <th style="border: 1px solid #FFFFFF; padding: 8px;">RICEF/CD #</th>
                        <th style="border: 1px solid #FFFFFF; padding: 8px;">Description</th>
                        <th style="border: 1px solid #FFFFFF; padding: 8px;">Functional Area</th>
                        <th style="border: 1px solid #FFFFFF; padding: 8px;">MR</th>
                    </tr>
                    <tr>
                        <td style="border: 1px solid #FFFFFF; padding: 8px;">{pwo}</td>
                        <td style="border: 1px solid #FFFFFF; padding: 8px;">{ricef_cd}</td>
                        <td style="border: 1px solid #FFFFFF; padding: 8px;">{description}</td>
                        <td style="border: 1px solid #FFFFFF; padding: 8px;">{functional_area}</td>
                        <td style="border: 1px solid #FFFFFF; padding: 8px;">{datetime.now().strftime("%B/%Y")}</td>
                    </tr>
                </table>            
            
            <p>Kindly let me know if you have conflicts with the proposed time so I can adjust depending on your availability.</p>            
            
            <p>Regards,</p>
            <p>Att,</p>

            <p><strong>{signature}</strong><br>
            Accenture Brazil ABAP Team</p>

            <p style="color: #00B050; font-size: 10pt;">Antes de imprimir, pense em sua responsabilidade com o MEIO AMBIENTE<br>
            <em>Before printing, think about your responsibility with the ENVIRONMENT</em></p>

            <p>Regards,</p>
            <p>Att,</p>
            </body>
            </html>
            """
        
        # Define o corpo do e-mail em HTML
        mail_item.HTMLBody = html_body

        # Obtém o diretório atual
        current_directory = os.getcwd()

        # Define o caminho para salvar o e-mail em formato .msg
        msg_path = os.path.join(current_directory, "meeting.msg")

        # Salva o e-mail em formato .msg
        mail_item.SaveAs(msg_path, 3)  # 3 corresponde ao formato .msg


    except Exception as e:
        print(f"Erro ao criar reunião: {e}")

# Exemplo de uso
subject = "Reunião de Equipe"
start_time = datetime(2024, 10, 10, 15, 0)  # 10 de outubro de 2024, às 15:00
end_time = datetime(2024, 10, 10, 16, 0)  # 10 de outubro de 2024, às 16:00
attendees = ["email1@dominio.com", "email2@dominio.com"]  # Substitua pelos emails dos participantes

create_teams_meeting(subject="Reunião de Equipe", 
                     emails="email1@dominio.com;email2@dominio.com;", 
                     pwo="PWOTEST", 
                     ricef_cd="12345/30009", 
                     description="TESTE DE REUNIAO",
                     functional_area="OTC",
                     signature="Diego Alves")
