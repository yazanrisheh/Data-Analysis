import streamlit as st
from pygwalker.api.streamlit import StreamlitRenderer
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
from dotenv import load_dotenv

load_dotenv()
# Adjust the width of the Streamlit page
st.set_page_config(page_title="Data Engineering", layout="wide")
st.title("Visualize your Data")

st.warning("Important: If your column names include Arabic text, you need to remove them first.")

# File uploader for CSV or Excel files
uploaded_file = st.file_uploader("Upload your CSV or Excel file", type=["csv", "xlsx"])

# Define caching function for memory efficiency
@st.cache_resource
def get_pyg_renderer(dataframe: pd.DataFrame) -> StreamlitRenderer:
    # If you want to use the feature of saving chart config, set `spec_io_mode="rw"`
    return StreamlitRenderer(dataframe)

if uploaded_file is not None:
    # Check if the uploaded file is an Excel file
    if uploaded_file.name.endswith('.xlsx'):
        # Load the Excel file and display the sheet selection option
        excel_file = pd.ExcelFile(uploaded_file)
        sheet_name = st.selectbox("Select the sheet", excel_file.sheet_names)
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
    else:
        # If it's a CSV file, read it directly
        df = pd.read_csv(uploaded_file)
    
    # Render the PyGWalker explorer
    renderer = get_pyg_renderer(df)
    renderer.explorer(key="explorer_tab")

    # Display the first 10 rows of the data in the sidebar
    st.sidebar.write("**First 10 rows of your data**")
    st.sidebar.write(df.head(10))
    st.sidebar.write("--------------------------------------------------------")
    st.sidebar.write(df.describe())

    # Generate a heatmap for correlation matrix of all columns
    st.subheader("Correlation Heatmap")
    st.info("Note this only applies to numerical only columns and not text")

    # Select only numeric columns for correlation
    corr = df.select_dtypes(include=['float64', 'int64']).corr()

    # Plot the heatmap using seaborn
    fig, ax = plt.subplots(figsize=(10, 8))
    sns.heatmap(corr, annot=True, cmap="coolwarm", linewidths=0.5, ax=ax)

    # Display the heatmap in Streamlit
    st.pyplot(fig)

else:
    st.write("Please upload a file to get started.")
