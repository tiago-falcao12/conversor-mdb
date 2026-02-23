import streamlit as st
import jpype
import jpype.imports
from jpype.types import *
import pandas as pd
import os
import glob
import re
import zipfile
import tempfile
from pathlib import Path

# -----------------------
# Função para path compatível com PyInstaller/Streamlit Cloud
# -----------------------
def resource_path(relative_path):
    return os.path.join(os.path.dirname(__file__), relative_path)

# -----------------------
# Inicia JVM com todos os jars na pasta jars/
# -----------------------
def start_jvm():
    if not jpype.isJVMStarted():
        jar_dir = resource_path("jars")
        jars = [os.path.join(jar_dir, jar) for jar in os.listdir(jar_dir) if jar.endswith(".jar")]
        if not jars:
            raise RuntimeError("Nenhum arquivo .jar encontrado na pasta jars/")
        jpype.startJVM(classpath=jars)

# -----------------------
# Sanitize filenames
# -----------------------
def sanitize_filename(name):
    return re.sub(r'[^\w\- ]', '_', str(name)).strip()

def convert_java_value(value):
    if value is None:
        return None
    return str(value)

# -----------------------
# Conversão MDB -> CSV
# -----------------------
def convert_mdb_to_csv(mdb_path, output_dir):
    from com.healthmarketscience.jackcess import DatabaseBuilder

    db = DatabaseBuilder.open(jpype.java.io.File(mdb_path))

    table_names = list(db.getTableNames())

    for table_name in table_names:
        table_name_py = str(table_name)

        try:
            table = db.getTable(table_name)
            columns = [str(col.getName()) for col in table.getColumns()]
            rows = []
            for row in table:
                row_data = [convert_java_value(row.get(col)) for col in columns]
                rows.append(row_data)
            df = pd.DataFrame(rows, columns=columns)
            safe_name = sanitize_filename(table_name_py)
            output_file = os.path.join(output_dir, f"{safe_name}.csv")
            df.to_csv(output_file, index=False, encoding="utf-8")
        except Exception as e:
            if "FileNotFoundException" in str(e):
                continue
            else:
                st.warning(f"Erro na tabela {table_name_py}: {e}")

    db.close()

# -----------------------
# Interface Streamlit
# -----------------------
st.set_page_config(page_title="Conversor MDB → CSV", layout="centered")
st.title("📂 Conversor MDB para CSV")

uploaded_file = st.file_uploader("Faça upload do arquivo .mdb", type=["mdb"])

if uploaded_file:
    with tempfile.TemporaryDirectory() as tmpdir:
        mdb_path = os.path.join(tmpdir, uploaded_file.name)
        with open(mdb_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        output_dir = os.path.join(tmpdir, "csv_output")
        os.makedirs(output_dir, exist_ok=True)

        with st.spinner("Processando arquivo..."):
            start_jvm()
            convert_mdb_to_csv(mdb_path, output_dir)

        # Criar zip com todos os CSVs
        zip_path = os.path.join(tmpdir, "resultado.zip")
        with zipfile.ZipFile(zip_path, "w") as zipf:
            for file in Path(output_dir).glob("*.csv"):
                zipf.write(file, arcname=file.name)

        st.success("Conversão finalizada!")

        with open(zip_path, "rb") as f:
            st.download_button(
                label="⬇️ Baixar CSVs (ZIP)",
                data=f,
                file_name="csv_convertidos.zip",
                mime="application/zip"
            )