#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Extrai dados do RAT OI CPE gerado com blindagem [[FIELD:key=value]]
e imprime a máscara de ENCERRAMENTO DE CPE.

Requisitos:
  - PyMuPDF (fitz)
  - (Opcional) pytesseract + tesseract-ocr se quiser fallback de OCR
Uso:
  python extract_oi_cpe_to_mask.py "RAT_OI_CPE_13456789.pdf"
"""

import sys, re, os
from typing import Dict, List
import fitz  # PyMuPDF

# --------- (opcional) OCR fallback ----------
try:
    import pytesseract
    from PIL import Image
    OCR_AVAILABLE = True
except Exception:
    OCR_AVAILABLE = False


FIELD_RX = re.compile(r"\[\[FIELD:([^=\]]+)=([^\]]*)\]\]")

def read_all_text(doc: fitz.Document) -> str:
    text_parts: List[str] = []
    for pno in range(doc.page_count):
        text_parts.append(doc[pno].get_text("text") or "")
    return "\n".join(text_parts)

def parse_fields_from_text(txt: str) -> Dict[str, str]:
    """
    Junta múltiplos FIELDs com a mesma chave na ordem em que aparecem.
    """
    fields = {}
    order_seen = []
    for m in FIELD_RX.finditer(txt):
        k = m.group(1).strip()
        v = (m.group(2) or "").strip()
        if k not in fields:
            fields[k] = v
            order_seen.append(k)
        else:
            # concatena com espaço
            if v:
                sep = " " if (fields[k] and not fields[k].endswith(" ")) else ""
                fields[k] = f"{fields[k]}{sep}{v}"
    return fields

def clean_value(s: str) -> str:
    """
    Remove underscores exagerados, espaços duplos, etc.
    """
    if not s:
        return ""
    s = s.replace("\u00a0", " ")  # NBSP
    # Muitas RATs têm linhas com ________ para preenchimento manual:
    s = re.sub(r"[_]{2,}", "", s)
    s = re.sub(r"\s{2,}", " ", s)
    return s.strip()

def build_mask(fields: Dict[str, str]) -> str:
    # Leitura segura dos campos
    chamado        = clean_value(fields.get("numero_chamado",""))
    cliente        = clean_value(fields.get("cliente",""))
    tecnico        = clean_value(fields.get("tecnico",""))
    cliente_ciente = clean_value(fields.get("cliente_ciente",""))
    contato        = clean_value(fields.get("contato",""))
    suporte_mam    = clean_value(fields.get("suporte_mam",""))
    teste_final    = clean_value(fields.get("teste_final","")).upper()  # S/N/NA
    produtivo      = clean_value(fields.get("produtivo",""))
    ba_num         = clean_value(fields.get("ba_num",""))
    motivo_imp     = clean_value(fields.get("motivo_improdutivo",""))
    obs            = clean_value(fields.get("observacoes",""))

    # Equipamento (primeiro item blindado)
    modelo         = clean_value(fields.get("equip_modelo",""))
    sn             = clean_value(fields.get("equip_sn",""))
    status         = clean_value(fields.get("equip_status",""))

    # Texto padrão para "TESTADO NA REDE GERENCIADA COM"
    if teste_final == "S":
        testado_msg = "Teste final realizado com o CPE do cliente conectado ao circuito, validação de camada 3 concluída."
    else:
        testado_msg = "Sem teste final com o CPE do cliente conectado ao circuito no momento do atendimento."

    # PRODUTIVO linha
    prod_line = produtivo
    if produtivo == "sim-com BA":
        if ba_num:
            prod_line = f"sim-com BA – BA {ba_num}"
        else:
            prod_line = "sim-com BA"
    elif produtivo == "não-improdutivo" and motivo_imp:
        # Mantemos só o PRODUTIVO limpo; motivo vai para descrição (se quiser)
        pass

    # “acompanhado pelo analista X” (gruda no PRODUTIVO como solicitado)
    if suporte_mam:
        prod_line = f"{prod_line} - acompanhado pelo analista {suporte_mam}"

    # DESCRIÇÃO (deixa OBS apenas; se quiser anexar motivo BA / improdutivo aqui, ajuste)
    descricao = obs

    # Montagem da máscara
    lines = [
        "###ENCERRAMENTO DE CPE###",
        "",
        "&&& MAMINFO",
        "",
        f"Nº DA RAT: {chamado}",
        f"CIRCUITO: {chamado}",
        f"MODELO DO CPE: {modelo}",
        f"Nº DE SÉRIE DO CPE: {sn}",
    ]
    # status do equipamento embaixo do serial (se houver)
    if status:
        lines.append(status)

    lines.extend([
        f"CIENTE NO LOCAL: SR(A) {cliente_ciente}",
        f"SUPORTE PELO ANALISTA: {suporte_mam}",
        f"REALIZADO PELO TÉCNICO: {tecnico}",
        f"FOI REALIZADO TESTE FINAL COM O EQUIPAMENTO DO CLIENTE? {'sim' if teste_final=='S' else 'não'}",
        f"TESTADO NA REDE GERENCIADA COM: {testado_msg}",
        f"CONTATO: {contato}",
        "CONFIGURAÇÕES EXECUTADAS:",
        f"PRODUTIVO: {prod_line}",
        f"DESCRIÇÃO: {descricao}",
        ""
    ])
    return "\n".join(lines)

# --------- (Opcional) OCR fallback para um ou outro campo ----------
def try_ocr_fallback(pdf_path: str, fields: Dict[str,str]) -> Dict[str,str]:
    """
    Se quiser, tente OCR em 1–2 páginas para garimpar algo que faltou.
    Mantém simples por padrão.
    """
    if not OCR_AVAILABLE:
        return fields

    try:
        doc = fitz.open(pdf_path)
    except Exception:
        return fields

    # Ex.: tenta só a primeira página
    try:
        page = doc[0]
        pix = page.get_pixmap(dpi=200)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        text = pytesseract.image_to_string(img, lang="por")
        # Exemplo besta: se não veio contato via FIELD, tenta OCR do padrão "Contato"
        if not fields.get("contato"):
            m = re.search(r"Contato[:\s]*([^\n\r]+)", text, re.I)
            if m:
                fields["contato"] = clean_value(m.group(1))
    except Exception:
        pass
    finally:
        try: doc.close()
        except: pass

    return fields

# ====================== CLI =========================
def main():
    if len(sys.argv) < 2:
        print("Uso: python extract_oi_cpe_to_mask.py <arquivo.pdf>")
        sys.exit(1)

    pdf_path = sys.argv[1]
    if not os.path.exists(pdf_path):
        print(f"Arquivo não encontrado: {pdf_path}")
        sys.exit(2)

    try:
        doc = fitz.open(pdf_path)
        full_text = read_all_text(doc)
        doc.close()
    except Exception as e:
        print(f"Falha ao abrir PDF: {e}")
        sys.exit(3)

    fields = parse_fields_from_text(full_text)

    # (Opcional) fallback OCR muito leve, só se algo essencial estiver vazio
    essentials = ["numero_chamado", "cliente_ciente", "tecnico", "contato"]
    if any(not fields.get(k) for k in essentials):
        fields = try_ocr_fallback(pdf_path, fields)

    mask = build_mask(fields)
    print(mask)

if __name__ == "__main__":
    main()
