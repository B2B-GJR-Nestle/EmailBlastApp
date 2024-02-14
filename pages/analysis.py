import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title='B2B Email Blast App',page_icon="📊")
st.title("📈B2B GJR Email Blast Analysis Tool")
hide_st = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            <style>
            """
st.markdown(hide_st, unsafe_allow_html=True)

def load_data(file):
    try:
        if file.name.endswith('.csv'):
            data = pd.read_csv(file)
        elif file.name.endswith(('.xls', '.xlsx')):
            data = pd.read_excel(file)
        else:
            st.error("Unsupported file format. Please upload a CSV or Excel file.")
            return None

        return data
    except UnicodeDecodeError:
        st.error("Unable to decode the file with UTF-8 encoding. Please try a different encoding.")
        return None
    except Exception as e:
        st.error(f"Error reading the file: {e}")
        return None

def pie_chart(data, column, title):
    fig = px.pie(data, names=column, title=title)
    st.plotly_chart(fig)

def product_histogram(data):
    fig = px.histogram(data, x='Product', color='Product', title='Product Histogram')
    st.plotly_chart(fig)

def main():
    st.title('Data Visualization App')

    uploaded_file = st.file_uploader("Upload a CSV or Excel file", type=["csv", "xlsx"])

    if uploaded_file is not None:
        data = load_data(uploaded_file)

        if data is not None:
            st.subheader('Data Preview:')
            st.write(data.head())

            st.subheader('📨Sent Rate')
            pie_chart(data, 'STATUS', 'Pie Chart of STATUS')

            st.subheader('📥Performance (Email Replied)')
            pie_chart(data, 'Replied', 'Pie Chart of Performance')

            st.subheader('📊Top Product Category by Proposal Sent')
            product_histogram(data)

if __name__ == "__main__":
    main()

st.sidebar.image("Nestle_Signature.png")
st.sidebar.write("""<p style='font-size: 14px;'>This Web-App is designed to facilitate B2B executive to email blasts for PT Nestlé Indonesia Customer made by <b>Nestlé Management Trainee 2023<b> (Nestea)</p>""", unsafe_allow_html=True)
st.sidebar.write("""<p style='font-size: 13px;'>For any inquiries, error handling, or assistance, please feel free to reach us through Email: <a href="mailto:Ananda.Cahyo@id.nestle.com">Ananda.Cahyo@id.nestle.com</a></p>""", unsafe_allow_html=True)
