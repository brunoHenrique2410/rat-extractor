import streamlit as st
from extract_oi_cpe_from_pdf import extract_from_pdf
import tempfile
import os

st.set_page_config(page_title="RAT OI CPE → Encerramento", layout="centered")
st.title("📄 Extrator de RAT OI CPE → Máscara de Encerramento")

st.markdown(
    "Envie o **PDF gerado pelo seu app de RAT OI CPE**. "
    "Eu extraio as informações e monto a máscara pronta para copiar."
)

pdf_file = st.file_uploader("📎 Envie o PDF preenchido", type=["pdf"])

if pdf_file is not None:
    # salva em arquivo temporário
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(pdf_file.read())
        tmp_path = tmp.name

    try:
        mask = extract_from_pdf(tmp_path)
        st.subheader("✅ Máscara de encerramento")
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
        try:
            os.remove(tmp_path)
        except Exception:
            pass

st.divider()
st.caption("Dica: se algum campo não vier, verifique se o PDF foi gerado pelo app padrão e se possui 2 páginas.")

