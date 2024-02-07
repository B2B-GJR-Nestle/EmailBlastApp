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

from py3o.renderserver.client import RenderServerClient
import base64

img = Image.open('Nestle_Logo.png')
st.set_page_config(page_title='B2B Email Blast App', page_icon=img)
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


def send_email(subject, body, to_address, attachment_path, gmail_user, gmail_password):
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


def convert_to_pdf(input_path, output_path):
    client = RenderServerClient()
    with open(input_path, "rb") as f:
        content = base64.b64encode(f.read()).decode("utf-8")
    pdf_content = client.render_template(content, "docx", "pdf")
    with open(output_path, "wb") as f:
        f.write(base64.b64decode(pdf_content))


def merge_and_send_emails(excel_data, gmail_user, gmail_password, template_path, body_text, feature, subject_email):
    output_directory = 'Promotion'
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)
    placeholder = st.empty()
    placeholder.dataframe(excel_data)
    for index, row in excel_data.iterrows():
        # Skip rows where the 'Email' column is empty
        if pd.isnull(row['Email']):
            continue

        merge_data = {
            'RecipientName': row['CP'],
            'Salutation': row['Salutation'],
            'CompanyName': row['Company Name'],
        }
        if feature == "Promotion":
            template = template_path
        else:
            product = row['Product']
            if product in template_dict:
                template = template_dict[product]
            else:
                st.error(f"No template found for product: {product}")
                continue
        output_filename_docx = f"{output_directory} Program Feeding {row['Company Name']}.docx"
        output_filename_pdf = f"{output_directory} Program Feeding {row['Company Name']}.pdf"
        generate_document(template, output_filename_docx, merge_data)
        convert_to_pdf(output_filename_docx, output_filename_pdf)  # Convert to PDF

        # Use the provided body_text or a default if none is provided
        email_subject = subject_email if subject_email else f"Proposal Penawaran Kerjasama PT Nestle Indonesia & {row['Company Name']}"
        email_body = body_text
        send_email(email_subject, email_body, row['Email'], output_filename_pdf, gmail_user, gmail_password)
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
    # excel_data = pd.read_excel("C:/Users/ASUS/Downloads/SalesProj/DATABASE.xlsx")  # Default path

# Select feature: Proposal or Promotion
st.write(f"## Blast Features")
feature = st.selectbox("Select Feature", ["Proposal", "Promotion"])

# Upload Word document template
template_path = None
if feature == "Promotion":
    st.write(f"## 🎁 Promotional File")
    template_path = st.file_uploader("Upload Promotion Template", type=["docx"])

# Upload Word document templates for Proposal feature
template_dict = {}
products = ["BearBrand", "Nescafe", "Milo"]
if feature == "Proposal":
    for product in products:
        if product == "BearBrand":
            st.write(f"## 🥛{product}")
        elif product == "Nescafe":
            st.write(f"## ☕ {product}")
        else:
            st.write(f"## 🍫 {product}")
        template_path = st.file_uploader(f"Upload {product} Template", type=["docx"])

        if template_path:
            # Save the uploaded template to a temporary file
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_template:
                temp_template.write(template_path.read())
                template_dict[product] = temp_template.name

# Input for email body text
default_body_text = """Salam,
Semoga Bapak/Ibu keadaan baik. Saya mewakili tim PT. Nestlé Indonesia dengan senang hati ingin berbicara tentang peluang kerjasama program feeding karyawan yang dapat memberikan nilai tambah bagi perusahaan Anda.
Sebagai salah satu perusahaan makanan dan minuman yang memiliki komitmen tinggi terhadap kualitas dan kesejahteraan, kami ingin menjalin kolaborasi dengan perusahaan Anda. Keunggulan kerjasama ini meliputi kontinuitas pasokan produk kami yang andal, serta diskon khusus sebagai bentuk apresiasi atas kerjasama yang baik.
Untuk informasi lebih lanjut seputar produk listing dan harga, Anda dapat menemukannya dalam dokumen yang saya lampirkan. Kami sangat terbuka untuk berdiskusi lebih lanjut atau menjawab pertanyaan yang mungkin Anda miliki.
Terima kasih banyak untuk waktu dan perhatiannya. Kami berharap dapat menjalin kerjasama yang baik dan saling menguntungkan.
Salam,
Bimo Agung Laksono
B2B Executive, Greater Jakarta Region - PT Nestlé Indonesia
Phone: +6287776162577
Mail : Bimoagung27@gmail.com
"""

body_text = st.text_area("Enter Email Body Text", default_body_text)

subject_email = st.text_input("Enter Email Subject", f"Proposal Penawaran Kerjasama PT Nestle Indonesia & {excel_data.iloc[0]['Company Name']}" if excel_file else "")

if st.button("Execute Mail Merge"):
    merge_and_send_emails(excel_data, "b2b.gjr.nestle@gmail.com", "alks kzuv wczc efch", template_path, body_text, feature, subject_email)
