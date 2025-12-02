import streamlit as st
import pandas as pd

@st.cache_data
def load_data():
    return pd.read_excel("Cores-Info-Database.xlsx", keep_default_na=False)

df = load_data()

st.title("Magnetic Core Selector")

# Core selector
selected_core = st.selectbox(
    "Select a magnetic core",
    df["Core Type"].unique()
)

# Filter selected row
core = df[df["Core Type"] == selected_core].iloc[0]

# Pretty display
st.subheader("Dimensions")

for column, value in core.items():
    st.write(f"**{column}:** {value}")