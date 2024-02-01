import streamlit as st
import pandas as pd
import os
from docxtpl import DocxTemplate
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import tempfile

def generate_document(template_path, output_path, data):
    doc = DocxTemplate(template_path)
    doc.render(data)
    doc.save(output_path)

def send_email(subject, body, to_address, attachment_path, gmail_user, gmail_password, output_update_function):
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = gmail_user
    msg['To'] = to_address

    # Attach the Word document
    with open(attachment_path, "rb") as attachment:
        part = MIMEApplication(attachment.read(), Name=os.path.basename(attachment_path))
        part['Content-Disposition'] = f'attachment; filename="{os.path.basename(attachment_path)}"'
        msg.attach(part)

    # Add the body text
    msg.attach(MIMEText(body, 'plain'))

    # Connect to Gmail SMTP server
    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(gmail_user, gmail_password)
        #output_update_function(f"Sent to {msg['To']}")
        server.sendmail(gmail_user, to_address, msg.as_string())

def update_excel_status(df, email, status):
    df.loc[df['Email'] == email, 'STATUS'] = status
    return df

def merge_and_send_emails(excel_data, gmail_user, gmail_password, template_folder, output_update_function):
    output_directory = 'output_emails/'
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    # Display the initial table after uploading the Excel file
    placeholder = st.empty()
    placeholder.dataframe(excel_data)

    for index, row in excel_data.iterrows():
        merge_data = {
            'RecipientName': row['CP'],
            'Salutation': row['Salutation'],
            'CompanyName': row['Company Name'],
        }

        # Extract product information from the "Product" column
        product = row['Product']

        # Determine the template based on the extracted product
        template_filename = f"Proposal_{product}.docx"
        template_path = os.path.join(template_folder, template_filename)

        output_filename = f"{output_directory}Proposal_{row['Company Name']}.docx"
        subject = f"Proposal Penawaran Kerjasama PT Nestle Indonesia & {row['Company Name']}"

        generate_document(template_path, output_filename, merge_data)

        # Body text for the email
        body_text = """
Salam,

Semoga Bapak/Ibu keadaan baik. Saya mewakili tim Nestlé Indonesia dan dengan senang hati ingin berbicara tentang peluang kerjasama program feeding karyawan yang dapat memberikan nilai tambah bagi perusahaan Anda.

Sebagai salah satu perusahaan makanan dan minuman yang memiliki komitmen tinggi terhadap kualitas dan kesejahteraan, kami ingin menjalin kolaborasi dengan perusahaan Anda. Keunggulan kerjasama ini meliputi kontinuitas pasokan produk kami yang andal, serta diskon khusus sebagai bentuk apresiasi atas kerjasama yang baik.

Untuk informasi lebih lanjut seputar produk listing dan harga, Anda dapat menemukannya dalam dokumen yang saya lampirkan. Kami sangat terbuka untuk berdiskusi lebih lanjut atau menjawab pertanyaan yang mungkin Anda miliki.

Terima kasih banyak untuk waktu dan perhatiannya. Kami berharap dapat menjalin kerjasama yang baik dan saling menguntungkan.

Salam,

Bimo
B2B Executive Greater Jakarta Region - Nestle Indonesia
        """

        # Send the email with the Word document as an attachment
        send_email(subject, body_text, row['Email'], output_filename, gmail_user, gmail_password, output_update_function)
        
        # Update Excel status to 'Sent' and update the DataFrame
        excel_data = update_excel_status(excel_data, row['Email'], 'Sent')
        
        # Display the updated DataFrame after each row
        placeholder.dataframe(excel_data)

# Streamlit app
st.title("📑Mail Merge Application")

# Upload Excel file
excel_file = st.file_uploader("Upload Excel File", type=["xlsx"])
if excel_file:
    excel_data = pd.read_excel(excel_file)
else:
    st.warning('Please Upload Your Database File', icon="⚠️")
    #excel_data = pd.read_excel("C:/Users/ASUS/Downloads/SalesProj/DATABASE.xlsx")  # Default path

# Select template folder
selected_folder_path = st.text_input("Enter Template Folder Path", "path/to/templates")

# Execute Mail Merge Button
if st.button("Execute Mail Merge"):
    merge_and_send_emails(excel_data, "b2b.gjr.nestle@gmail.com", "alks kzuv wczc efch", selected_folder_path, st.empty())
