import streamlit as st
import pandas as pd
import plotly.express as px

def load_data(file):
    try:
        data = pd.read_csv(file, encoding='utf-8')  # Specify the encoding used in your CSV file
        return data
    except UnicodeDecodeError:
        st.error("Unable to decode the file with UTF-8 encoding. Please try a different encoding.")
        return None

def pie_chart(data):
    fig = px.pie(data, names='STATUS', title='Pie Chart of STATUS')
    st.plotly_chart(fig)

def product_histogram(data):
    fig = px.histogram(data, x='Product', color='Product', title='Product Histogram')
    st.plotly_chart(fig)

def main():
    st.title('Data Visualization App')

    uploaded_file = st.file_uploader("Upload a CSV or Excel file", type=["csv", "xlsx"])

    if uploaded_file is not None:
        data = load_data(uploaded_file)

        st.subheader('Data Preview:')
        st.write(data.head())

        st.subheader('Visualization 1: Pie Chart of STATUS')
        pie_chart(data)

        st.subheader('Visualization 2: Product Histogram')
        product_histogram(data)

if __name__ == "__main__":
    main()
