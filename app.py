import streamlit as st
from extract_oi_cpe_from_pdf import extract_from_pdf
import tempfile, os

st.set_page_config(page_title="RAT OI CPE → Encerramento", layout="centered")
st.title("📄 Extrator de RAT OI CPE → Máscara de Encerramento (limpa)")

st.markdown(
    "Envie o **PDF gerado pelo app de RAT OI CPE**. "
    "Eu extraio os dados, limpo sublinhados/aspas e monto a máscara **sem aspas**, só com os valores."
)

pdf_file = st.file_uploader("📎 Envie o PDF preenchido", type=["pdf"])

if pdf_file is not None:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(pdf_file.read())
        tmp_path = tmp.name

    try:
        mask = extract_from_pdf(tmp_path)
        st.subheader("✅ Máscara de encerramento (limpa)")
        st.code(mask, language="text")

        st.download_button(
            "⬇️ Baixar máscara (.txt)",
            data=mask,
            file_name="encerramento_cpe.txt",
            mime="text/plain"
        )
    except Exception as e:
        st.error("Falha ao processar o PDF.")
        st.exception(e)
    finally:
        try: os.remove(tmp_path)
        except: pass

st.caption("Se algum campo não vier, confira se o PDF é o do template do app (2 páginas).")
