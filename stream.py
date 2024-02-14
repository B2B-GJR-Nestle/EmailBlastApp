import streamlit as st
import pandas as pd
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from PIL import Image

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

def send_email(subject, body, to_address, attachment_data, attachment_filename, gmail_user, gmail_password):
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = gmail_user
    msg['To'] = to_address

    part = MIMEApplication(attachment_data, Name=attachment_filename)
    part['Content-Disposition'] = f'attachment; filename="{attachment_filename}"'
    msg.attach(part)

    msg.attach(MIMEText(body, 'plain'))

    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(gmail_user, gmail_password)
        server.sendmail(gmail_user, to_address, msg.as_string())

def merge_and_send_emails(excel_data, gmail_user, gmail_password, template_path, body_text, subject_text, feature):
    output_directory = 'Promotion'
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    for index, row in excel_data.iterrows():
        # Skip rows where the 'Email' column is empty
        if pd.isnull(row['Email']):
            continue

        if feature == "Promotion":
            if template_path:
                attachment_data = template_path.getvalue()
                attachment_filename = template_path.name
                subject = subject_text.format(company_name=row['Company Name'])
                email_body = body_text.format(CompanyName=row['Company Name'])
                send_email(subject, email_body, row['Email'], attachment_data, attachment_filename, gmail_user, gmail_password)
            else:
                st.error("No attachment uploaded for this email.")
                continue
        else:
            # For Proposal feature, generate document and attach
            merge_data = {
                'RecipientName': row['CP'],
                'Salutation': row['Salutation'],
                'CompanyName': row['Company Name'],
            }
            product = row['Product']
            if product in template_dict:
                template = template_dict[product]
            else:
                st.error(f"No template found for product: {product}")
                continue
            output_filename = f"{output_directory} Program Feeding {row['Company Name']}.docx"
            subject = subject_text.format(company_name=row['Company Name'])
            generate_document(template, output_filename, merge_data)
            email_body = body_text.format(CompanyName=row['Company Name'])
            send_email(subject, email_body, row['Email'], output_filename, gmail_user, gmail_password)

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

# Select feature: Proposal or Promotion
st.write(f"## Blast Features")
feature = st.selectbox("Select Feature", ["Proposal", "Promotion"])

# Upload Word document template
template_path = None
if feature == "Promotion":
    st.write(f"## 🎁 Promotional File")
    template_path = st.file_uploader("Upload Promotion Template", type=["docx","png", "jpg", "jpeg", "pdf"])

# Upload Word document templates for Proposal feature
template_dict = {}
products = ["General","BearBrand", "Nescafe"]
if feature == "Proposal":
    for product in products:
        if product == "General":    
            st.write(f"## 🍫 {product}")
        elif product == "BearBrand":    
            st.write(f"## 🥛{product}")
        else:
            st.write(f"## ☕ {product}")
        template_path = st.file_uploader(f"Upload {product} Template", type=["docx","pdf"])
        if template_path:
            # Save the uploaded template to a temporary file
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_template:
                temp_template.write(template_path.read())
                template_dict[product] = temp_template.name

# Input for email subject text
default_subject = "Proposal Penawaran Kerjasama PT Nestle Indonesia & {company_name}"
subject_text = st.text_input("Enter Email Subject", default_subject)

# Input for email body text
default_body = """Yth. Bapak/Ibu
{CompanyName},
di Tempat

Semoga Bapak/Ibu keadaan baik. Saya mewakili tim PT. Nestlé Indonesia dengan senang hati ingin berbicara tentang peluang kerjasama program feeding karyawan yang dapat memberikan nilai tambah bagi perusahaan Anda.

Sebagai salah satu perusahaan makanan dan minuman yang memiliki komitmen tinggi terhadap kualitas dan kesejahteraan, kami ingin menjalin kolaborasi dengan perusahaan Anda. Keunggulan kerjasama ini meliputi kontinuitas pasokan produk kami yang andal, serta diskon khusus sebagai bentuk apresiasi atas kerjasama yang baik.

Untuk informasi lebih lanjut seputar produk listing dan harga, Anda dapat menemukannya dalam dokumen yang saya lampirkan. Kami sangat terbuka untuk berdiskusi lebih lanjut atau menjawab pertanyaan yang mungkin Anda miliki.

Terima kasih banyak untuk waktu dan perhatiannya. Kami berharap dapat menjalin kerjasama yang baik dan saling menguntungkan.

Salam,

Bimo Agung Laksono
B2B Executive, Greater Jakarta Region - PT Nestlé Indonesia
Phone: +6287776162577 | Mail : Bimoagung27@gmail.com"""

body_text = st.text_area("Enter Email Body Text", default_body, height=300)

if st.button("Execute Mail Merge"):
    merge_and_send_emails(excel_data, "b2b.gjr.nestle@gmail.com", "alks kzuv wczc efch", template_path, body_text, subject_text, feature)
st.sidebar.image("Nestle_Signature.png")
st.sidebar.write("This Web-App is designed to facilitate B2B email blasts for PT Nestlé")
