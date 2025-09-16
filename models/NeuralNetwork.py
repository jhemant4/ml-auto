import streamlit as st
from sklearn.neural_network import MLPClassifier


def nn_param_selector(prefix="nn"):
    number_hidden_layers = st.number_input(
        "Number of hidden layers",
        1, 5, 1,
        key=f"{prefix}_n_layers"   # ✅ unique key
    )

    hidden_layer_sizes = []

    for i in range(number_hidden_layers):
        n_neurons = st.number_input(
            f"Neurons in layer {i+1}",
            2, 200, 100, 25,
            key=f"{prefix}_layer_{i+1}"   # ✅ unique key per layer
        )
        hidden_layer_sizes.append(n_neurons)

    hidden_layer_sizes = tuple(hidden_layer_sizes)

    params = {"hidden_layer_sizes": hidden_layer_sizes}

    # ✅ always return a model
    model = MLPClassifier(**params)
    return model
