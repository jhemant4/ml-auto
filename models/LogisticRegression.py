import numpy as np
import streamlit as st
from sklearn.linear_model import LogisticRegression

def lr_param_selector(prefix="lr"):
    solver = st.selectbox(
        "Solver",
        options=["lbfgs", "newton-cg", "liblinear", "sag", "saga"],
        key=f"{prefix}_solver"
    )

    if solver in ["newton-cg", "lbfgs", "sag"]:
        penalties = ["l2", "none"]
    elif solver == "saga":
        penalties = ["l1", "l2", "none", "elasticnet"]
    elif solver == "liblinear":
        penalties = ["l1"]

    penalty = st.selectbox("Penalty", options=penalties, key=f"{prefix}_penalty")

    C = st.number_input("Inverse Regularization Strength (C)", 0.1, 2.0, 1.0, 0.1, key=f"{prefix}_c")
    C = np.round(C, 3)

    max_iter = st.number_input("Max Iterations", 100, 2000, step=50, value=100, key=f"{prefix}_max_iter")

    return LogisticRegression(solver=solver, penalty=penalty, C=C, max_iter=max_iter)
