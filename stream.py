import streamlit as st
import pandas as pd
import os
import requests
from docxtpl import DocxTemplate
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import tempfile

def generate_document(template, output_path, data):
    doc = DocxTemplate(template)
    doc.render(data)
    doc.save(output_path)

def download_template_from_github(repo_url, template_path):
    template_url = f"{repo_url}/blob/main/{template_path}"
    response = requests.get(template_url)
    if response.status_code == 200:
        template_content = response.content
        return template_content
    else:
        st.error(f"Failed to download template from GitHub. Please check the URL: {template_url}")

def send_email(subject, body, to_address, attachment_path, gmail_user, gmail_password, output_update_function):
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = gmail_user
    msg['To'] = to_address

    with open(attachment_path, "rb") as attachment:
        part = MIMEApplication(attachment.read(), Name=os.path.basename(attachment_path))
        part['Content-Disposition'] = f'attachment; filename="{os.path.basename(attachment_path)}"'
        msg.attach(part)

    msg.attach(MIMEText(body, 'plain'))

    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(gmail_user, gmail_password)
        server.sendmail(gmail_user, to_address, msg.as_string())

def update_excel_status(df, email, status):
    df.loc[df['Email'] == email, 'STATUS'] = status
    return df

def merge_and_send_emails(excel_data, gmail_user, gmail_password, template_dict, output_update_function):
    output_directory = 'output_emails'
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    placeholder = st.empty()
    placeholder.dataframe(excel_data)

    for index, row in excel_data.iterrows():
        merge_data = {
            'RecipientName': row['CP'],
            'Salutation': row['Salutation'],
            'CompanyName': row['Company Name'],
        }

        product = row['Product']

        if product in template_dict:
            template = template_dict[product]
            output_filename = f"{output_directory}Proposal_{row['Company Name']}_{product}.docx"
            subject = f"Proposal Penawaran Kerjasama PT Nestle Indonesia & {row['Company Name']} ({product})"
            generate_document(template, output_filename, merge_data)

            body_text = """
Salam,
HAI
"""

            send_email(subject, body_text, row['Email'], output_filename, gmail_user, gmail_password, output_update_function)

            excel_data = update_excel_status(excel_data, row['Email'], 'Sent')
            placeholder.dataframe(excel_data)

# Streamlit app
st.title("📑Mail Merge Application")

excel_file = st.file_uploader("Upload Excel File", type=["xlsx"])
if excel_file:
    excel_data = pd.read_excel(excel_file)
else:
    st.warning('Please Upload Your Database File', icon="⚠️")
    #excel_data = pd.read_excel("C:/Users/ASUS/Downloads/SalesProj/DATABASE.xlsx")  # Default path

github_repo_url = st.text_input("Enter GitHub Repository URL", "https://github.com/B2B-GJR-Nestle/EmailBlastApp")

template_dict = {}
products = ["BearBrand", "Nescafe", "Milo"]

for product in products:
    st.write(f"## {product} Template")
    template_path = st.file_uploader(f"Upload {product} Template", type=["docx"])
    
    if template_path:
        template_content = template_path.read()
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_template:
            temp_template.write(template_content)
            template_dict[product] = temp_template.name

if st.button("Execute Mail Merge"):
    merge_and_send_emails(excel_data, "b2b.gjr.nestle@gmail.com", "alks kzuv wczc efch", template_dict, st.empty())
