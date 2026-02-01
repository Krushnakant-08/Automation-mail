import pandas as pd
from pptx import Presentation
import smtplib
import ssl
import shutil
import win32com.client
import os
from email.message import EmailMessage
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
APP_PASSWORD = os.getenv("APP_PASSWORD")
SMTP_SERVER = os.getenv("SMTP_SERVER")
SMTP_PORT = 25 #465
ssl_context = ssl.create_default_context()

# Read data and set template
df = pd.read_excel("Sample.xlsx")
template_path = "./student.pptx"

# Create output directories if they don't exist
os.makedirs("./Files/pptx", exist_ok=True)
os.makedirs("./Files/pdfs", exist_ok=True)

# Initialize PowerPoint once for all conversions
powerpoint = win32com.client.Dispatch("PowerPoint.Application")
powerpoint.Visible = 1

for index, row in df.iterrows():
    
    # Create a copy of the template
    prs = Presentation(template_path)
    
    # Replace placeholders in all slides
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                if "{{Name}}" in shape.text:
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.text = run.text.replace("{{Name}}", row['Name'])
                if "{{Event}}" in shape.text:
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.text = run.text.replace("{{Event}}", row['Event'])
    
    output_path = f"./Files/pptx/{row['Name']}_{row['Event']}.pptx"
    prs.save(output_path)
    print(f"Saved: {output_path}")
    
    # Convert to PDF using PowerPoint COM automation
    pdf_output_path = f"./Files/pdfs/{row['Name']}_{row['Event']}.pdf"
    # Convert to absolute paths
    abs_input_path = os.path.abspath(output_path)
    abs_output_path = os.path.abspath(pdf_output_path)
    
    presentation = powerpoint.Presentations.Open(abs_input_path, WithWindow=False)
    presentation.SaveAs(abs_output_path, 32)  # 32 = ppSaveAsPDF
    presentation.Close()
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
                             filename=f"{row['Name']}_{row['Event']}.pdf")
        
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, context=ssl_context) as server:
            server.login(SENDER_EMAIL, APP_PASSWORD)
            server.send_message(msg)
            print(f"-> Successfully sent email with PDF attachment to: {row['Email']}")
            
    except Exception as e:
        print(f"-> Failed to send email to {row['Email']}. Error: {e}")

# Close PowerPoint application
powerpoint.Quit()

print("\nAll documents have been processed and emails sent.")

