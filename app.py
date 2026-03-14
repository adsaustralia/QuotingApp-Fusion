
import streamlit as st
import pandas as pd
import numpy as np
import math

st.title("SQM Quoting Tool")

# -----------------------------
# Example price table
# -----------------------------
rates = {
    "Ferrous": 12.0,
    "Banner": 19.5,
    "Corflute": 16.0
}

# -----------------------------
# SQM CALCULATION
# -----------------------------
def sqm_calc(width_mm, height_mm):
    if pd.isna(width_mm) or pd.isna(height_mm):
        return np.nan
    return (width_mm/1000) * (height_mm/1000)

# -----------------------------
# SQM MARKUP RULE
# -----------------------------
def sqm_markup_factor(sqm_each):
    if pd.isna(sqm_each):
        return 1.0
    
    sqm_each = float(sqm_each)

    if 0 <= sqm_each <= 1:
        return 3.5
    elif 1 < sqm_each <= 3:
        return 1.8
    elif 3 < sqm_each <= 5:
        return 1.35
    else:
        return 1.0

# -----------------------------
# UI
# -----------------------------
width = st.number_input("Width (mm)", value=240)
height = st.number_input("Height (mm)", value=193)
qty = st.number_input("Quantity", value=1)
material = st.selectbox("Material", list(rates.keys()))
ds = st.checkbox("Double sided")

ds_loading_pct = st.number_input("DS loading (%)", value=20.0)/100

# -----------------------------
# CALCULATIONS
# -----------------------------
sqm_each = sqm_calc(width, height)
sqm_total = sqm_each * qty

sqm_rate = rates.get(material, 0)

ds_factor = 1 + ds_loading_pct if ds else 1
markup = sqm_markup_factor(sqm_each)

price = sqm_total * sqm_rate * ds_factor * markup

# -----------------------------
# REVIEW
# -----------------------------
review = pd.DataFrame({
    "width_mm":[width],
    "height_mm":[height],
    "sqm_each":[sqm_each],
    "qty":[qty],
    "total_sqm":[sqm_total],
    "rate":[sqm_rate],
    "ds_factor":[ds_factor],
    "sqm_markup_factor":[markup],
    "price":[price]
})

st.subheader("Quote Review")
st.dataframe(review)

st.success(f"Final Price: ${price:.2f}")
