import streamlit as st
import pandas as pd
import plotly.express as px

def load_data(file):
    try:
        if file.name.endswith('.csv'):
            data = pd.read_csv(file, encoding='utf-8')
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

        if data is not None:
            st.subheader('Data Preview:')
            st.write(data.head())

            # Create two columns for side-by-side visualizations
            col1, col2 = st.columns(2)

            with col1:
                st.subheader('Visualization 1: Pie Chart of STATUS')
                pie_chart(data, 'STATUS', 'Pie Chart of STATUS')

            with col2:
                st.subheader('Visualization 2: Pie Chart of Performance')
                pie_chart(data, 'Performance', 'Pie Chart of Performance')

            st.subheader('Visualization 3: Product Histogram')
            product_histogram(data)

if __name__ == "__main__":
    main()
