import pandas as pd
from docxtpl import DocxTemplate
from docx2pdf import convert
import smtplib
import ssl
from email.message import EmailMessage
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
APP_PASSWORD = os.getenv("APP_PASSWORD")
SMTP_SERVER = os.getenv("SMTP_SERVER")
SMTP_PORT = 465
ssl_context = ssl.create_default_context()

# Read data and set template
df = pd.read_excel("Sample.xlsx")
template_path = "./student.docx"


for index, row in df.iterrows():
    
    doc_context = {
        'Name': row['Name'],
        'Email': row['Email']
    }
    
   
    doc = DocxTemplate(template_path)
    doc.render(doc_context)
    
    output_path = f"./Files/docx/{row['Name']}.docx"
    doc.save(output_path)
    print(f"Saved: {output_path}")
    
    # Convert to PDF
    pdf_output_path = f"./Files/pdfs/{row['Name']}.pdf"
    convert(output_path, pdf_output_path)
    print(f"Exported to PDF: {pdf_output_path}")
    
    # Send email
    try:
        msg = EmailMessage()
        msg['Subject'] = "This is a trial mail"
        msg['From'] = SENDER_EMAIL
        msg['To'] = row['Email']
        
        body = f"""
        Hi {row['Name']},
        This is a trial email with your document attached.
        """
        msg.set_content(body.strip())
        
    
        with open(pdf_output_path, 'rb') as f:
            pdf_data = f.read()
            msg.add_attachment(pdf_data, maintype='application', 
                             subtype='pdf', 
                             filename=f"{row['Name']}.pdf")
        
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, context=ssl_context) as server:
            server.login(SENDER_EMAIL, APP_PASSWORD)
            server.send_message(msg)
            print(f"-> Successfully sent email with PDF attachment to: {row['Email']}")
            
    except Exception as e:
        print(f"-> Failed to send email to {row['Email']}. Error: {e}")

print("\nAll documents have been processed and emails sent.")

