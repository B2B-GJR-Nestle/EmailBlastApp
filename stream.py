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
from PIL import Image

img = Image.open('Nestle_Logo.png')
st.set_page_config(page_title='B2B Email Blast App',page_icon=img)
st.title("📑B2B GJR Email Blast Application")
hide_st = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            <style>
            """
st.markdown(hide_st, unsafe_allow_html=True)


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

def merge_and_send_emails(excel_data, gmail_user, gmail_password, template_dict, body_text, output_update_function):
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

            # Use the provided body_text or a default if none is provided
            email_body = body_text
            send_email(subject, email_body, row['Email'], output_filename, gmail_user, gmail_password, output_update_function)

            excel_data = update_excel_status(excel_data, row['Email'], 'Sent')
            placeholder.dataframe(excel_data)

# Streamlit app
# Upload Excel or CSV file
excel_file = st.file_uploader("Upload Excel/CSV File", type=["xlsx", "csv"])
if excel_file:
    # Use pandas to read either Excel or CSV based on file extension
    if excel_file.name.endswith('.xlsx'):
        excel_data = pd.read_excel(excel_file)
    elif excel_file.name.endswith('.csv'):
        excel_data = pd.read_csv(excel_file)
else:
    st.warning('Please Upload Your Database File', icon="⚠️")
    #excel_data = pd.read_excel("C:/Users/ASUS/Downloads/SalesProj/DATABASE.xlsx")  # Default path

github_repo_url = st.text_input("Enter GitHub Repository URL", "https://github.com/B2B-GJR-Nestle/EmailBlastApp")

template_dict = {}
products = ["BearBrand", "Nescafe", "Milo"]

for product in products:
    if product == "BearBrand":    
        st.write(f"## 🥛{product}")
    elif product == "Nescafe":    
        st.write(f"## ☕ {product}")
    else:
        st.write(f"## 🍫 {product}")
    template_path = st.file_uploader(f"Upload {product} Template", type=["docx"])
    
    if template_path:
        template_content = template_path.read()
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_template:
            temp_template.write(template_content)
            template_dict[product] = temp_template.name

# Input for email body text
default = """
Salam,

Semoga Bapak/Ibu keadaan baik. Saya mewakili tim Nestlé Indonesia dan dengan senang hati ingin berbicara tentang peluang kerjasama program feeding karyawan yang dapat memberikan nilai tambah bagi perusahaan Anda.

Sebagai salah satu perusahaan makanan dan minuman yang memiliki komitmen tinggi terhadap kualitas dan kesejahteraan, kami ingin menjalin kolaborasi dengan perusahaan Anda. Keunggulan kerjasama ini meliputi kontinuitas pasokan produk kami yang andal, serta diskon khusus sebagai bentuk apresiasi atas kerjasama yang baik.

Untuk informasi lebih lanjut seputar produk listing dan harga, Anda dapat menemukannya dalam dokumen yang saya lampirkan. Kami sangat terbuka untuk berdiskusi lebih lanjut atau menjawab pertanyaan yang mungkin Anda miliki.

Terima kasih banyak untuk waktu dan perhatiannya. Kami berharap dapat menjalin kerjasama yang baik dan saling menguntungkan.

Salam,

Bimo
B2B Executive Greater Jakarta Region - Nestle Indonesia
"""
body_text = st.text_area("Enter Email Body Text", default)

if st.button("Execute Mail Merge"):
    merge_and_send_emails(excel_data, "b2b.gjr.nestle@gmail.com", "alks kzuv wczc efch", template_dict, body_text, st.empty())
