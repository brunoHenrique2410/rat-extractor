# -*- coding: utf-8 -*-
"""
Extrai campos blindados [[FIELD:key=value]] do PDF e monta a máscara de encerramento.

Função principal: extract_from_pdf(pdf_bytes) -> (mask_text: str, fields: dict)
"""

import re
from typing import Dict, Tuple
import fitz  # PyMuPDF

# ---------- regex dos fields ----------
FIELD_RX = re.compile(r"\[\[FIELD:([^=\]]+)=([^\]]*)\]\]")

def _read_all_text(doc: fitz.Document) -> str:
    parts = []
    for i in range(doc.page_count):
        parts.append(doc[i].get_text("text") or "")
    return "\n".join(parts)

def _parse_fields(txt: str) -> Dict[str, str]:
    """Concatena valores de mesmos fields na ordem em que aparecem."""
    out: Dict[str, str] = {}
    for m in FIELD_RX.finditer(txt):
        k = (m.group(1) or "").strip()
        v = (m.group(2) or "").strip()
        if not k:
            continue
        if k not in out:
            out[k] = v
        else:
            if v:
                sep = "" if (not out[k] or out[k].endswith(" ")) else " "
                out[k] = f"{out[k]}{sep}{v}"
    return out

def _clean(s: str) -> str:
    if not s:
        return ""
    s = s.replace("\u00a0", " ")  # NBSP
    # remove longas sequências de underscores/traços usados como linha
    s = re.sub(r"[_]{2,}", "", s)
    s = re.sub(r"\s{2,}", " ", s)
    return s.strip()

def _build_mask(fields: Dict[str, str]) -> str:
    chamado        = _clean(fields.get("numero_chamado",""))
    cliente        = _clean(fields.get("cliente",""))
    tecnico        = _clean(fields.get("tecnico",""))
    cliente_ciente = _clean(fields.get("cliente_ciente",""))
    contato        = _clean(fields.get("contato",""))
    suporte_mam    = _clean(fields.get("suporte_mam",""))
    teste_final    = _clean(fields.get("teste_final","")).upper()  # S/N/NA
    produtivo      = _clean(fields.get("produtivo",""))
    ba_num         = _clean(fields.get("ba_num",""))
    motivo_imp     = _clean(fields.get("motivo_improdutivo",""))
    obs            = _clean(fields.get("observacoes",""))

    # Equipamento (primeiro item)
    modelo         = _clean(fields.get("equip_modelo",""))
    sn             = _clean(fields.get("equip_sn",""))
    status         = _clean(fields.get("equip_status",""))

    # Frase padrão do “TESTADO …”
    if teste_final == "S":
        testado_msg = "Teste final realizado com o CPE do cliente conectado ao circuito, validação de camada 3 concluída."
        teste_simnao = "sim"
    else:
        testado_msg = "Sem teste final com o CPE do cliente conectado ao circuito no momento do atendimento."
        teste_simnao = "não"

    # PRODUTIVO line
    prod_line = produtivo
    if produtivo == "sim-com BA":
        prod_line = f"sim-com BA" + (f" – BA {ba_num}" if ba_num else "")
    # “acompanhado pelo analista …”
    if suporte_mam:
        prod_line = f"{prod_line} - acompanhado pelo analista {suporte_mam}"

    # Descrição deixa somente OBS (personalize se quiser incluir BA/motivo)
    descricao = obs

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
    if status:
        lines.append(status)

    lines.extend([
        f"CLIENTE NO LOCAL: SR(A) {cliente_ciente}",
        f"SUPORTE PELO ANALISTA: {suporte_mam}",
        f"REALIZADO PELO TÉCNICO: {tecnico}",
        f"FOI REALIZADO TESTE FINAL COM O EQUIPAMENTO DO CLIENTE? {teste_simnao}",
        f"TESTADO NA REDE GERENCIADA COM: {testado_msg}",
        f"CONTATO: {contato}",
        "CONFIGURAÇÕES EXECUTADAS:",
        f"PRODUTIVO: {prod_line}",
        f"DESCRIÇÃO: {descricao}",
        "",
    ])
    return "\n".join(lines)

def extract_from_pdf(pdf_bytes: bytes) -> Tuple[str, Dict[str,str]]:
    """
    Lê pdf_bytes, extrai [[FIELD:...=...]] e retorna (mask_text, fields_dict).
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        txt = _read_all_text(doc)
    finally:
        doc.close()

    fields = _parse_fields(txt)
    mask = _build_mask(fields)
    return mask, fields
