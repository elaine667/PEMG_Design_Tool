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

# Get single row of data
core = df[df["Core Type"] == selected_core].iloc[0]

st.subheader("Core Dimensions")

# Display dimensions
for column, value in core.items():
    st.write(f"**{column}:** {value}")

# ---- Calculate Core Area Product ----

# Try to convert to float, skip if "Not Available"
try:
    Ae = float(core["A_c,e [mm^2]"])
    Aw = float(core["A_w [mm^2]"])
    area_product = Ae * Aw
    result_text = f"{area_product:.2f} mm⁴"  # nice formatting
except:
    result_text = "Not Available"

st.markdown("---")
st.subheader("Calculated Value")

st.write(f"**Core Area Product:** {result_text}")