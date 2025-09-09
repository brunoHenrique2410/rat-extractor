# extract_oi_cpe_from_pdf.py
# Extrai dados da RAT OI CPE gerada pelo seu app e imprime a máscara ###ENCERRAMENTO DE CPE###
# Limpeza avançada: remove "Bilhete_____"/sublinhados/aspas, separa "Produtivo" do "acompanhado pelo analista",
# e monta a saída SEM aspas e apenas com os valores.
#
# Requisitos: PyMuPDF (fitz)

import re, fitz
from typing import List, Tuple, Optional, Dict

# ---------- utils básicos ----------
def words(page) -> List[Tuple[float,float,float,float,str,int,int,int]]:
    return page.get_text("words")

def search_first(page, labels) -> Optional[fitz.Rect]:
    if isinstance(labels, str):
        labels = [labels]
    for t in labels:
        try:
            hits = page.search_for(t)
            if hits:
                return hits[0]
        except Exception:
            pass
    return None

def text_in_rect(page, rect: fitz.Rect) -> str:
    ws = [w for w in words(page) if fitz.Rect(w[0],w[1],w[2],w[3]).intersects(rect)]
    ws.sort(key=lambda w: (w[1], w[0]))
    return " ".join(w[4] for w in ws).strip()

def grab_right_of(page, label_variants, dx=6, dy=1, w=420, h=24) -> str:
    r = search_first(page, label_variants)
    if not r:
        return ""
    rect = fitz.Rect(r.x1 + dx, r.y0 + dy, r.x1 + dx + w, r.y0 + dy + h)
    return text_in_rect(page, rect)

def nearest_word(page, x, y, max_dist=10) -> Optional[str]:
    cand = []
    for w in words(page):
        cx = (w[0]+w[2])/2; cy = (w[1]+w[3])/2
        d2 = (cx-x)**2 + (cy-y)**2
        if d2 <= max_dist**2:
            cand.append((d2, w[4]))
    if not cand:
        return None
    cand.sort(key=lambda t: t[0])
    return cand[0][1]

# ---------- limpeza de valores ----------
UNDERLINE_PAT = re.compile(r"[_]{2,}")          # ________
DUP_SPACES    = re.compile(r"\s{2,}")
QUOTES_PAT    = re.compile(r'^[\'"]+|[\'"]+$')

def clean_value(s: str) -> str:
    """Remove aspas, sublinhados, lixo como 'Bilhete___________', e normaliza espaços."""
    if not s:
        return ""
    s = s.strip()
    s = QUOTES_PAT.sub("", s)
    s = s.replace("\u00A0", " ")  # NBSP -> espaço normal
    # Remove tokens 'Bilhete' e similares
    s = re.sub(r"\b[Bb]ilhete\b", "", s)
    # Remove 'Contato' quando vier grudado ao número
    s = re.sub(r"\b[Cc]ontato\b", "", s)
    # Remove sequência de sublinhados
    s = UNDERLINE_PAT.sub(" ", s)
    # Colapsa espaços
    s = DUP_SPACES.sub(" ", s).strip(" -:;\t")
    return s

def extract_first_digits(s: str) -> str:
    """Pega o primeiro bloco de dígitos (ex.: '13456789 Bilhete____' -> '13456789')."""
    if not s:
        return ""
    m = re.search(r"\b(\d{4,})\b", s)
    return m.group(1) if m else clean_value(s)

def only_digits_or_clean(s: str) -> str:
    """Se houver bloco numérico longo (contato), prioriza; senão devolve limpo."""
    s2 = clean_value(s)
    m = re.search(r"\b(\d{5,})\b", s2)
    return m.group(1) if m else s2

# ---------- extrações específicas ----------
def extract_numero_chamado(page1) -> str:
    for labels in [["Número do Bilhete","Numero do Bilhete"],
                   ["Designação do Circuito","Designacao do Circuito"]]:
        raw = grab_right_of(page1, labels, dx=8, dy=1, w=420, h=24)
        val = extract_first_digits(raw)
        if val:
            return val
    return ""

def extract_identificacao_pagina1(page1) -> Dict[str,str]:
    out = {
        "tecnico": "",
        "cliente_ciente": "",
        "contato": "",
        "aceitacao": "",
        "teste_final": ""  # "S" / "N" / "NA"
    }
    out["tecnico"] = clean_value(grab_right_of(page1, ["Técnico","Tecnico"], dx=8, dy=1))
    out["cliente_ciente"] = clean_value(grab_right_of(page1, ["Cliente Ciente","Cliente  Ciente"], dx=8, dy=1))
    out["contato"] = only_digits_or_clean(grab_right_of(page1, ["Contato"], dx=8, dy=1))
    out["aceitacao"] = clean_value(grab_right_of(page1, ["Aceitação do serviço pelo responsável",
                                                         "Aceitacao do servico pelo responsavel"], dx=8, dy=1))

    # Detecta “X” em S / N / N/A
    wan_label = search_first(page1, ["Teste de conectividade WAN","Teste final com equipamento do cliente"])
    if wan_label:
        pos_S  = wan_label.x1 + 138
        pos_N  = wan_label.x1 + 165
        pos_NA = wan_label.x1 + 207
        ymark  = wan_label.y0 + 11
        hit_S  = (nearest_word(page1, pos_S,  ymark) or "").upper()
        hit_N  = (nearest_word(page1, pos_N,  ymark) or "").upper()
        hit_NA = (nearest_word(page1, pos_NA, ymark) or "").upper()
        if hit_S == "X":
            out["teste_final"] = "S"
        elif hit_N == "X":
            out["teste_final"] = "N"
        elif hit_NA == "X":
            out["teste_final"] = "NA"
    return out

def extract_observacoes(page2) -> str:
    r = search_first(page2, ["OBSERVAÇÕES","Observacoes","Observações"])
    if not r:
        return ""
    rect = fitz.Rect(r.x0, r.y1 + 20, r.x0 + 540, r.y1 + 20 + 220)
    return clean_value(text_in_rect(page2, rect))

def extract_problema(page2) -> str:
    r = search_first(page2, ["PROBLEMA ENCONTRADO","Problema Encontrado"])
    if not r:
        return ""
    rect = fitz.Rect(r.x0, r.y1 + 20, r.x0 + 540, r.y1 + 20 + 220)
    return clean_value(text_in_rect(page2, rect))

def extract_acao(page2) -> str:
    r = search_first(page2, ["AÇÃO CORRETIVA","Acao Corretiva","Ação Corretiva"])
    if not r:
        return ""
    rect = fitz.Rect(r.x0, r.y1 + 20, r.x0 + 540, r.y1 + 20 + 160)
    return clean_value(text_in_rect(page2, rect))

def extract_equipamento_principal(page2) -> Dict[str,str]:
    """
    Lê as linhas da seção 'EQUIPAMENTOS NO CLIENTE' (formato impresso pelo app)
    e retorna o PRIMEIRO item: tipo / numero_serie / modelo / status
    """
    out = {"tipo":"","N° de Serie":"","modelo":"","status":""}
    title = search_first(page2, ["EQUIPAMENTOS NO CLIENTE","Equipamentos no Cliente"])
    if not title:
        return out

    base_x = title.x0
    TOP_OFFSET = 36
    ROW_DY     = 26

    for row_idx in range(0, 12):
        y = title.y1 + TOP_OFFSET + row_idx * ROW_DY
        band = fitz.Rect(base_x, y-4, base_x+560, y+18)
        line = text_in_rect(page2, band)
        if not line:
            continue
        # "Tipo: X | S/N: Y | Mod: Z | Status: W"
        def pick(tag):
            m = re.search(tag + r"\s*:\s*([^|]+)", line, flags=re.I)
            return clean_value(m.group(1)) if m else ""
        out = {
            "tipo": pick(r"(?:Tipo)"),
            "N° de Serie": pick(r"(?:S/N)"),
            "modelo": pick(r"(?:Mod)"),
            "status": pick(r"(?:Status)"),
        }
        if any(out.values()):
            break
    return out

# ---------- montagem da máscara (sem aspas) ----------
def build_mask(
    numero_chamado: str,
    equip: Dict[str,str],
    tecnico: str,
    cliente_ciente: str,
    contato: str,
    suporte_mam_hint: str,
    teste_final: str,
    observacoes_raw: str,
    problema: str,
    acao: str,
) -> str:
    # Extrai produtivo / suporte / BA das áreas de texto
    obs = observacoes_raw or ""

    # Produtivo
    m_prod = re.search(r"\bProdutivo:\s*([^\n\r]+)", obs, flags=re.I)
    produtivo = clean_value(m_prod.group(1)) if m_prod else ""

    # Suporte pelo analista
    m_sup = re.search(r"acompanhado pelo analista\s+([A-Za-zÀ-ÿ\s]+)", obs, flags=re.I)
    suporte_mam = clean_value(m_sup.group(1)) if m_sup else clean_value(suporte_mam_hint)

    # BA (se houver)
    m_ba = re.search(r"\bBA:\s*([A-Za-z0-9\-_/]+)", acao or "", flags=re.I)
    ba_num = clean_value(m_ba.group(1)) if m_ba else ""

    # Limpa a descrição: remove a linha "Produtivo: ..." para não duplicar
    obs_clean = re.sub(r"(?im)^.*\bProdutivo:\s*[^\n\r]+$", "", obs).strip()

    # teste final -> sim/não
    tf = (teste_final or "").upper()
    tf_answer = "sim" if tf == "S" else ("não" if tf == "N" else "não")

    # Mensagem padrão de rede gerenciada:
    testado_com = (
        "Teste final realizado com o CPE do cliente conectado ao circuito, validação de camada 3 concluída."
        if tf_answer == "sim"
        else "Sem teste final com o equipamento do cliente no momento do atendimento."
    )

    # PRODUTIVO: se "sim-com BA" e houver BA -> inclui no próprio valor
    if produtivo.lower().startswith("sim-com ba") and ba_num:
        produtivo_fmt = f"sim-com BA ({ba_num})"
    else:
        produtivo_fmt = produtivo

    modelo = clean_value(equip.get("modelo") or "")
    serial = clean_value(equip.get("numero_serie") or "")
    status = clean_value(equip.get("status") or "")
    tecnico = clean_value(tecnico)
    cliente_ciente = clean_value(cliente_ciente)
    contato = only_digits_or_clean(contato)
    numero_chamado = extract_first_digits(numero_chamado)

    # Monta a máscara SEM aspas
    lines = []
    lines.append("###ENCERRAMENTO DE CPE###")
    lines.append("")
    lines.append("&&& MAMINFO")
    lines.append("")
    lines.append(f"Nº DA RAT: {numero_chamado}")
    lines.append(f"CIRCUITO: {numero_chamado}")
    lines.append(f"MODELO DO CPE: {modelo}")
    lines.append(f"Nº DE SÉRIE DO CPE: {serial}")
    if status:
        lines.append(f"Status do equipamento: {status}")
    lines.append(f"CIENTE NO LOCAL: SR(A) {cliente_ciente}")
    lines.append(f"SUPORTE PELO ANALISTA: {suporte_mam}")
    lines.append(f"REALIZADO PELO TÉCNICO: {tecnico}")
    lines.append(f"FOI REALIZADO TESTE FINAL COM O EQUIPAMENTO DO CLIENTE? {tf_answer}")
    lines.append(f"TESTADO NA REDE GERENCIADA COM: {testado_com}")
    lines.append(f"CONTATO: {contato}")
    lines.append("CONFIGURAÇÕES EXECUTADAS:")
    lines.append(f"PRODUTIVO: {produtivo_fmt}")
    lines.append(f'DESCRIÇÃO: {obs_clean}')
    return "\n".join(lines)

# ---------- API pública ----------
def extract_from_pdf(path: str) -> str:
    doc = fitz.open(path)
    try:
        page1 = doc[0]
        page2 = doc[1] if doc.page_count >= 2 else doc[0]

        numero = extract_numero_chamado(page1)
        ident  = extract_identificacao_pagina1(page1)
        equip  = extract_equipamento_principal(page2)
        obs    = extract_observacoes(page2)
        prob   = extract_problema(page2)
        acao   = extract_acao(page2)

        mask = build_mask(
            numero_chamado=numero,
            equip=equip,
            tecnico=ident.get("tecnico",""),
            cliente_ciente=ident.get("cliente_ciente",""),
            contato=ident.get("contato",""),
            suporte_mam_hint="",  # pode vir das observações
            teste_final=ident.get("teste_final",""),
            observacoes_raw=obs,
            problema=prob,
            acao=acao,
        )
        return mask
    finally:
        doc.close()
