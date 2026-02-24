import streamlit as st
import subprocess
import tempfile
import os
import zipfile

st.set_page_config(page_title="Conversor MDB para CSV", layout="centered")
st.title("Conversor de Arquivos Microsoft Access (.mdb) para CSV")
st.markdown("Faça upload de um arquivo .mdb e obtenha as tabelas em CSV.")

uploaded_file = st.file_uploader("Escolha um arquivo .mdb", type=["mdb"])

if uploaded_file is not None:
    # Verifica tamanho (limite de 200MB)
    if uploaded_file.size > 200 * 1024 * 1024:
        st.error("Arquivo muito grande. O limite é 200 MB.")
        st.stop()

    # Salva o arquivo temporariamente
    with tempfile.NamedTemporaryFile(delete=False, suffix=".mdb") as tmp_file:
        tmp_file.write(uploaded_file.getbuffer())
        mdb_path = tmp_file.name

    try:
        # Lista tabelas com mdb-tables
        result = subprocess.run(
            ["mdb-tables", "-1", mdb_path],
            capture_output=True,
            text=True,
            check=True,
            timeout=30
        )
        tables = [t.strip() for t in result.stdout.split("\n") if t.strip()]
    except subprocess.TimeoutExpired:
        st.error("Tempo limite excedido ao ler o arquivo.")
        st.stop()
    except subprocess.CalledProcessError as e:
        st.error(f"Erro ao ler o arquivo .mdb: {e.stderr}")
        st.stop()
    except Exception as e:
        st.error(f"Erro inesperado: {str(e)}")
        st.stop()

    if not tables:
        st.warning("Nenhuma tabela encontrada no arquivo.")
        st.stop()

    st.success(f"Arquivo carregado! Encontradas {len(tables)} tabelas.")

    selected_tables = st.multiselect(
        "Selecione as tabelas para converter (deixe vazio caso não queira converter nenhuma)",
        options=tables,
        default=tables
    )
    if not selected_tables:
        selected_tables = tables

    if st.button("Converter para CSV"):
        with st.spinner("Convertendo..."):
            with tempfile.TemporaryDirectory() as output_dir:
                csv_files = []
                progress_bar = st.progress(0)

                for i, table in enumerate(selected_tables):
                    safe_name = "".join(c for c in table if c.isalnum() or c in (' ', '_')).rstrip()
                    csv_path = os.path.join(output_dir, f"{safe_name}.csv")

                    try:
                        with open(csv_path, "w") as f:
                            subprocess.run(
                                ["mdb-export", mdb_path, table],
                                stdout=f,
                                stderr=subprocess.PIPE,
                                check=True,
                                text=True,
                                timeout=60
                            )
                        csv_files.append(csv_path)
                    except subprocess.TimeoutExpired:
                        st.warning(f"Tempo limite excedido na tabela '{table}'.")
                    except subprocess.CalledProcessError as e:
                        st.warning(f"Erro na tabela '{table}': {e.stderr}")
                    except Exception as e:
                        st.warning(f"Erro inesperado na tabela '{table}': {str(e)}")

                    progress_bar.progress((i + 1) / len(selected_tables))

                if not csv_files:
                    st.error("Nenhuma tabela foi convertida.")
                else:
                    zip_path = os.path.join(tempfile.gettempdir(), "converted.zip")
                    with zipfile.ZipFile(zip_path, "w") as zipf:
                        for file in csv_files:
                            zipf.write(file, arcname=os.path.basename(file))

                    with open(zip_path, "rb") as f:
                        st.download_button(
                            label="Baixar arquivo ZIP com os CSVs",
                            data=f,
                            file_name="tabelas_convertidas.zip",
                            mime="application/zip"
                        )

    # Remove o arquivo .mdb temporário
    os.unlink(mdb_path)