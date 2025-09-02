import streamlit as st
import pythoncom
import win32com.client
import tempfile
import pandas as pd
import pywintypes
import json
import plotly.graph_objects as go

st.set_page_config(page_title="Extraer nodos y barras", layout="wide")
st.title("📐 Extraer nodos y barras de Robot + JSON + Visualización 3D")

uploaded_file = st.file_uploader("Sube tu archivo de Robot (.rtd)", type=["rtd"])

def iter_labels_from_getall(coll):
    """Intenta obtener etiquetas con GetAll() sin usar Count: iteramos hasta que falle."""
    labels = []
    try:
        arr = coll.GetAll()
        if arr:
            i = 1
            while True:
                try:
                    lbl = arr.Get(i)          # devuelve la etiqueta/numero real
                    labels.append(int(lbl))
                    i += 1
                except Exception:
                    break
    except Exception:
        pass
    return labels

def fallback_iter_labels_by_get(coll, max_tries=10000):
    """Fallback: probar Nodes.Get(i) con índices crecientes hasta que falle muchas veces consecutivas."""
    labels = []
    i = 1
    consecutive_failures = 0
    while i <= max_tries and consecutive_failures < 20:
        try:
            item = coll.Get(i)
            try:
                lab = int(getattr(item, "Number", getattr(item, "Label", i)))
            except Exception:
                lab = i
            labels.append(int(lab))
            consecutive_failures = 0
            i += 1
        except Exception:
            consecutive_failures += 1
            i += 1
    return sorted(list(set(labels)))

def extract_nodes_and_bars(project):
    """Extrae nodos y barras sin usar .Count."""
    structure = project.Structure
    if not structure:
        raise RuntimeError("No se pudo acceder a project.Structure")

    # NODOS
    nodes_coll = structure.Nodes
    labels = iter_labels_from_getall(nodes_coll)
    if not labels:
        labels = fallback_iter_labels_by_get(nodes_coll)

    nodes_data = []
    for lbl in labels:
        try:
            n = nodes_coll.Get(lbl)
            nodes_data.append({
                "id": int(getattr(n, "Number", lbl)),
                "x": getattr(n, "X", None),
                "y": getattr(n, "Y", None),
                "z": getattr(n, "Z", None)
            })
        except Exception:
            continue

    # BARRAS
    bars_coll = structure.Bars
    bar_labels = []
    try:
        bar_labels = iter_labels_from_getall(bars_coll)
        if not bar_labels:
            bar_labels = fallback_iter_labels_by_get(bars_coll)
    except Exception:
        bar_labels = []

    bars_data = []
    for lbl in bar_labels:
        try:
            b = bars_coll.Get(lbl)
            bars_data.append({
                "id": int(getattr(b, "Number", lbl)),
                "start_node": int(getattr(b, "StartNode", None)),
                "end_node": int(getattr(b, "EndNode", None)),
                "section": getattr(b, "SectionName", "")
            })
        except Exception:
            continue

    df_nodes = pd.DataFrame(nodes_data).sort_values("id").reset_index(drop=True)
    df_bars = pd.DataFrame(bars_data).sort_values("id").reset_index(drop=True)
    return df_nodes, df_bars

if uploaded_file:
    suffix = ".rtd"
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    try:
        tmp.write(uploaded_file.getbuffer())
        tmp.flush()
    finally:
        tmp.close()
    rtd_path = tmp.name
    st.success(f"📂 Archivo temporal guardado en: {rtd_path}")

    if st.button("🔎Extraer Información"):
        resultado_container = st.empty()
        try:
            pythoncom.CoInitialize()
            resultado_container.info("Conectando con Robot...")
            try:
                robot = win32com.client.GetActiveObject("Robot.Application")
            except Exception:
                robot = win32com.client.Dispatch("Robot.Application")
            robot.Visible = False

            resultado_container.info("Analizando proyecto en Robot...")
            proj = robot.Project
            proj.Open(rtd_path)

            resultado_container.info("Extrayendo nodos y barras...")
            df_nodes, df_bars = extract_nodes_and_bars(proj)

            # Crear JSON con toda la info
            data = {
                "nodes": df_nodes.to_dict(orient="records"),
                "bars": df_bars.to_dict(orient="records")
            }
            json_str = json.dumps(data, indent=4)
            st.download_button("⬇️ Descargar estructura (JSON)", json_str, file_name="estructura_robot.json", mime="application/json")

            # Mostrar tablas en la app
            if not df_nodes.empty:
                st.subheader("📍 Nodos")
                st.dataframe(df_nodes, use_container_width=True)
            else:
                st.warning("No se extrajeron nodos.")

            if not df_bars.empty:
                st.subheader("🔗 Barras")
                st.dataframe(df_bars, use_container_width=True)
            else:
                st.info("No se encontraron barras.")

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

                fig.update_layout(
                    scene=dict(
                        xaxis_title="X",
                        yaxis_title="Y",
                        zaxis_title="Z"
                    ),
                    title="Visualización 3D de la estructura"
                )
                st.plotly_chart(fig, use_container_width=True)

            st.success("✅ Extracción y visualización completadas.")

        except pywintypes.com_error as ce:
            st.error(f"Error COM: {ce}")
        except Exception as e:
            st.error(f"Error inesperado: {e}")
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
