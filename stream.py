import streamlit as st
import pandas as pd
import os
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
        server.sendmail(gmail_user, to_address, msg.as_string())

def update_excel_status(df, email, status):
    df.loc[df['Email'] == email, 'STATUS'] = status
    return df

def merge_and_send_emails(excel_data, gmail_user, gmail_password, template_dict, output_update_function):
    output_directory = 'output_emails'
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
        if product in template_dict:
            template = template_dict[product]
            output_filename = f"{output_directory}Proposal_{row['Company Name']}_{product}.docx"
            subject = f"Proposal Penawaran Kerjasama PT Nestle Indonesia & {row['Company Name']} ({product})"
            generate_document(template, output_filename, merge_data)

            # Body text for the email
            body_text = """
Salam,
# ... (Your email body)
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

# Upload Word templates
template_dict = {}
products = ["BearBrand", "Nescafe", "Milo"]

for product in products:
    template_file = st.file_uploader(f"Upload Template for {product}", type=["docx"])
    if template_file:
        template_dict[product] = template_file.name

# Execute Mail Merge Button
if st.button("Execute Mail Merge"):
    merge_and_send_emails(excel_data, "b2b.gjr.nestle@gmail.com", "alks kzuv wczc efch", template_dict, st.empty())
