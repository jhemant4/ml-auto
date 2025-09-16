from sklearn.naive_bayes import GaussianNB
import streamlit as st

def nb_param_selector(prefix="nb"):
    st.text("No parameters to tune for Naive Bayes")
    return GaussianNB()   # ✅ always return a model
