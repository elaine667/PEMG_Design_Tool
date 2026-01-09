import streamlit as st
import pandas as pd

@st.cache_data
def load_data():
    return pd.read_excel("Cores-Info-Database.xlsx", keep_default_na=False)

df = load_data()

st.title("Magnetic Core Selector")

# Core selector
core_name = st.selectbox("Select a magnetic core", df["Core Type"].unique())
core = df[df["Core Type"] == core_name].iloc[0]

# Safe float conversion
def safe_float(x):
    try:
        return float(x)
    except:
        return None


st.markdown("---")
st.subheader("Core Dimensions")

# Map dataframe columns → LaTeX variable names
dim_map = {
    "A_c,e [mm^2]": r"A_e \; (\mathrm{mm}^2)",
    "A_w [mm^2]": r"A_w \; (\mathrm{mm}^2)",
    "A_c,min [mm^2|": r"A_{c,\min} \; (\mathrm{mm}^2)",
    "l_e [mm]": r"l_e \; (\mathrm{mm})",
    "l_t [mm]": r"l_t \; (\mathrm{mm})",
    "Bobbin window area": r"A_{b,w} \; (\mathrm{mm}^2)",
    "l_N": r"l_N"
}

# Display each dimension with LaTeX formatting
for col, latex_var in dim_map.items():
    val = core.get(col, "Not Available")
    st.latex(latex_var + " = " + str(val))


# -----------------------------------------------------------
# Calculations
# -----------------------------------------------------------
st.markdown("---")
st.subheader("Calculated Values")

Ae = safe_float(core.get("A_c,e [mm^2]"))
Aw = safe_float(core.get("A_w [mm^2]"))
Ac_min = safe_float(core.get("A_c,min [mm^2|"))
lt = safe_float(core.get("l_t [mm]"))

# --------- AREA PRODUCT (example) ---------
if Ae is not None and Aw is not None:
    Ap = Ae * Aw
    st.latex(r"A_p = A_e \cdot A_w")
    st.latex(r"A_p = " + f"{Ap:.3f}" + r"\; \mathrm{mm}^4")
else:
    st.latex(r"A_p = \text{Not Available}")

# --------- Kg parameter ---------
if Ac_min is not None and Aw is not None and lt not in (None, 0):
    Kg = Ac_min * Aw / lt
    st.latex(r"K_g = \frac{A_{c,\min} \cdot A_w}{l_t}")
    st.latex(r"K_g = " + f"{Kg:.4f}" + r"\; \mathrm{mm}^3")
else:
    st.latex(r"K_g = \frac{A_{c,\min} \cdot A_w}{l_t}")
    st.latex(r"K_g = \text{Not Available}")

# ===========================================================
# NEW ADDITION: Picture display
# ===========================================================
st.markdown("---")
st.subheader("Picture:")

image_path = core.get("Image Path", "")

if image_path:
    st.image(
        image_path,
        use_container_width=True
    )
else:
    st.text("No picture available for this core.")