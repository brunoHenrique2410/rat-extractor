# -*- coding: utf-8 -*-
import io
import json
import streamlit as st

from extract_oi_cpe_from_pdf import extract_from_pdf

st.set_page_config(page_title="Extrator RAT OI CPE → Máscara", layout="centered")
st.title("🧾 Extrator de RAT OI CPE → Máscara de Encerramento")

st.caption("Faça upload do PDF gerado (com blindagem [[FIELD:...]] na última página).")

up = st.file_uploader("PDF da RAT OI CPE", type=["pdf"])

col1, col2 = st.columns([1,1])

if up is not None:
    pdf_bytes = up.read()
    try:
        mask, fields = extract_from_pdf(pdf_bytes)

        st.success("PDF lido com sucesso!")

        st.subheader("🧩 Máscara gerada")
        st.code(mask, language="markdown")

        st.download_button(
            "⬇️ Baixar máscara (.txt)",
            data=mask.encode("utf-8"),
            file_name="encerramento_cpe.txt",
            mime="text/plain"
        )

        with col1:
            st.subheader("Campos detectados")
            st.json(fields, expanded=False)

        with col2:
            st.subheader("Copiar rapidamente")
            st.text_area("Copiar/editar", value=mask, height=240)

    except Exception as e:
        st.error("Falha ao processar o PDF.")
        st.exception(e)
else:
    st.info("Envie um PDF para extrair os dados e gerar a máscara.")
