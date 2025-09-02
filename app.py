import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import json

st.set_page_config(page_title="Visualizar estructura", layout="wide")
st.title("📐 Visualización de nodos y barras")

uploaded_json = st.file_uploader("Sube un archivo JSON de estructura", type=["json"])

if uploaded_json:
    data = json.load(uploaded_json)
    df_nodes = pd.DataFrame(data["nodes"])
    df_bars = pd.DataFrame(data["bars"])

    st.subheader("📍 Nodos")
    st.dataframe(df_nodes)

    st.subheader("🔗 Barras")
    st.dataframe(df_bars)

    # Gráfico 3D
    if not df_nodes.empty:
        fig = go.Figure()

        # Nodos
        fig.add_trace(go.Scatter3d(
            x=df_nodes["x"], y=df_nodes["y"], z=df_nodes["z"],
            mode="markers+text",
            text=df_nodes["id"],
            marker=dict(size=5, color="red"),
            name="Nodos"
        ))

        # Barras
        for _, bar in df_bars.iterrows():
            try:
                n1 = df_nodes[df_nodes["id"] == bar["start_node"]].iloc[0]
                n2 = df_nodes[df_nodes["id"] == bar["end_node"]].iloc[0]
                fig.add_trace(go.Scatter3d(
                    x=[n1["x"], n2["x"]],
                    y=[n1["y"], n2["y"]],
                    z=[n1["z"], n2["z"]],
                    mode="lines",
                    line=dict(color="blue", width=3),
                    name=f"Barra {bar['id']}"
                ))
            except IndexError:
                continue

        fig.update_layout(
            scene=dict(xaxis_title="X", yaxis_title="Y", zaxis_title="Z"),
            title="Visualización 3D de la estructura"
        )
        st.plotly_chart(fig, use_container_width=True)
