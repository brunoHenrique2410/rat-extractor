# extract_oi_cpe_from_pdf.py
# Extrai dados da RAT OI CPE e imprime a máscara ###ENCERRAMENTO DE CPE###
# Ajuste: serial não confunde com "EQUIPAMENTOS" (filtro de candidatos + prioridade S/N)

import re, fitz
from typing import List, Tuple, Optional, Dict

# -------- utils básicos --------
def words(page) -> List[Tuple[float,float,float,float,str,int,int,int]]:
    return page.get_text("words")

def search_first(page, labels) -> Optional[fitz.Rect]:
    if isinstance(labels, str): labels=[labels]
    for t in labels:
        try:
            hits = page.search_for(t)
            if hits: return hits[0]
        except Exception:
            pass
    return None

def search_all(page, labels) -> List[fitz.Rect]:
    if isinstance(labels, str): labels=[labels]
    out=[]
    for t in labels:
        try: out.extend(page.search_for(t))
        except Exception: pass
    return out

def text_in_rect(page, rect: fitz.Rect) -> str:
    ws = [w for w in words(page) if fitz.Rect(w[0],w[1],w[2],w[3]).intersects(rect)]
    ws.sort(key=lambda w: (w[1], w[0]))
    return " ".join(w[4] for w in ws).strip()

def grab_right_of(page, label_variants, dx=6, dy=1, w=420, h=24) -> str:
    r = search_first(page, label_variants)
    if not r: return ""
    rect = fitz.Rect(r.x1 + dx, r.y0 + dy, r.x1 + dx + w, r.y0 + dy + h)
    return text_in_rect(page, rect)

def nearest_word(page, x, y, max_dist=10) -> Optional[str]:
    cand=[]
    for w in words(page):
        cx=(w[0]+w[2])/2; cy=(w[1]+w[3])/2
        d2=(cx-x)**2+(cy-y)**2
        if d2<=max_dist**2: cand.append((d2,w[4]))
    if not cand: return None
    cand.sort(key=lambda t:t[0])
    return cand[0][1]

# -------- limpeza --------
UNDERLINE_PAT = re.compile(r"[_]{2,}")
DUP_SPACES    = re.compile(r"\s{2,}")
QUOTES_PAT    = re.compile(r'^[\'"]+|[\'"]+$')

def clean_value(s: str) -> str:
    if not s: return ""
    s = s.strip()
    s = QUOTES_PAT.sub("", s)
    s = s.replace("\u00A0"," ")
    s = re.sub(r"\b[Bb]ilhete\b","", s)
    s = re.sub(r"\b[Cc]ontato\b","", s)
    s = UNDERLINE_PAT.sub(" ", s)
    s = DUP_SPACES.sub(" ", s).strip(" -:;")
    return s

def extract_first_digits(s: str) -> str:
    if not s: return ""
    m = re.search(r"\b(\d{4,})\b", s)
    return m.group(1) if m else clean_value(s)

def only_digits_or_clean(s: str) -> str:
    s2 = clean_value(s)
    m = re.search(r"\b(\d{5,})\b", s2)
    return m.group(1) if m else s2

# -------- extrações específicas --------
def extract_numero_chamado(page1) -> str:
    for labels in [["Número do Bilhete","Numero do Bilhete"],
                   ["Designação do Circuito","Designacao do Circuito"]]:
        raw = grab_right_of(page1, labels, dx=8, dy=1, w=420, h=24)
        val = extract_first_digits(raw)
        if val: return val
    return ""

def extract_identificacao_pagina1(page1) -> Dict[str,str]:
    out = {"tecnico":"", "cliente_ciente":"", "contato":"", "aceitacao":"", "teste_final":""}
    out["tecnico"] = clean_value(grab_right_of(page1, ["Técnico","Tecnico"], dx=8, dy=1))
    out["cliente_ciente"] = clean_value(grab_right_of(page1, ["Cliente Ciente","Cliente  Ciente"], dx=8, dy=1))
    out["contato"] = only_digits_or_clean(grab_right_of(page1, ["Contato"], dx=8, dy=1))
    out["aceitacao"] = clean_value(grab_right_of(page1,
        ["Aceitação do serviço pelo responsável","Aceitacao do servico pelo responsavel"], dx=8, dy=1))

    wan_label = search_first(page1, ["Teste de conectividade WAN","Teste final com equipamento do cliente"])
    if wan_label:
        pos_S  = wan_label.x1 + 138
        pos_N  = wan_label.x1 + 165
        pos_NA = wan_label.x1 + 207
        ymark  = wan_label.y0 + 11
        hit_S  = (nearest_word(page1, pos_S,  ymark) or "").upper()
        hit_N  = (nearest_word(page1, pos_N,  ymark) or "").upper()
        hit_NA = (nearest_word(page1, pos_NA, ymark) or "").upper()
        if hit_S == "X": out["teste_final"]="S"
        elif hit_N == "X": out["teste_final"]="N"
        elif hit_NA == "X": out["teste_final"]="NA"
    return out

def extract_observacoes(page) -> str:
    r = search_first(page, ["OBSERVAÇÕES","Observacoes","Observações"])
    if not r: return ""
    rect = fitz.Rect(r.x0, r.y1+20, r.x0+540, r.y1+20+240)
    return clean_value(text_in_rect(page, rect))

def extract_problema(page) -> str:
    r = search_first(page, ["PROBLEMA ENCONTRADO","Problema Encontrado"])
    if not r: return ""
    rect = fitz.Rect(r.x0, r.y1+20, r.x0+540, r.y1+20+240)
    return clean_value(text_in_rect(page, rect))

def extract_acao(page) -> str:
    r = search_first(page, ["AÇÃO CORRETIVA","Acao Corretiva","Ação Corretiva"])
    if not r: return ""
    rect = fitz.Rect(r.x0, r.y1+20, r.x0+540, r.y1+20+160)
    return clean_value(text_in_rect(page, rect))

# ---- equipamento (robusto) ----
MODEL_KEYWORDS = ["synway","aligera","syn way","syn-way"]
STATUS_CHOICES = [
    "equipamento no local","instalado pelo técnico","retirado pelo técnico",
    "spare técnico","técnico não levou equipamento",
    # variações sem acento
    "instalado pelo tecnico","retirado pelo tecnico","spare tecnico","tecnico nao levou equipamento"
]
# Palavras que NUNCA podem virar serial
SN_STOPWORDS = {
    "EQUIPAMENTOS","NO","CLIENTE","EQUIPAMENTOSNOCLIENTE","STATUS","TIPO","MOD","MODELO","ITEM",
    "PROBLEMA","OBSERVAÇÕES","OBSERVACOES","ACAO","AÇÃO","CORRETIVA","S","N","NA"
}

def _looks_like_serial(token: str) -> bool:
    """Serial precisa ter pelo menos 6 chars e conter ALGUM dígito; e não ser stopword."""
    if not token: return False
    t = token.strip().upper()
    if len(t) < 6: return False
    if t in SN_STOPWORDS: return False
    if not re.search(r"\d", t):  # precisa ter dígito
        return False
    # evita tokens com muitos separadores
    if t.count("|")>0: return False
    return True

def _extract_equipment_block_from_page(page) -> Dict[str,str]:
    """
    1) Tenta linha única "Tipo: ... | S/N: ... | Mod: ... | Status: ..."
    2) Se não achar, faz fallback por regex com filtros (serial com dígito etc).
    """
    out = {"tipo":"","numero_serie":"","modelo":"","status":""}
    anchors = search_all(page, ["EQUIPAMENTOS NO CLIENTE","Equipamentos no Cliente"])
    if not anchors:
        return out

    for anc in anchors:
        # 1) Linha única (scan linhas virtuais)
        base_x = anc.x0
        TOP_OFFSET = 36
        ROW_DY     = 26
        for row_idx in range(0, 14):
            y = anc.y1 + TOP_OFFSET + row_idx*ROW_DY
            band = fitz.Rect(base_x, y-4, base_x+560, y+18)
            line = text_in_rect(page, band)
            if not line: continue

            def pick(tag):
                m = re.search(tag + r"\s*:\s*([^|]+)", line, flags=re.I)
                return clean_value(m.group(1)) if m else ""
            tipo   = pick(r"(?:Tipo)")
            s_n    = pick(r"(?:S/?N)")
            modelo = pick(r"(?:Mod|Modelo)")
            status = pick(r"(?:Status)")

            if any([tipo, s_n, modelo, status]):
                # valida o serial
                if s_n and not _looks_like_serial(s_n):
                    s_n = ""
                out["tipo"]=tipo
                out["numero_serie"]=s_n
                out["modelo"]=modelo
                out["status"]=status
                # se encontrou um serial válido, retorna; senão continua tentando
                if out["numero_serie"] or any([modelo,status,tipo]):
                    return out

        # 2) Fallback: coletar tokens na área ampla
        area = fitz.Rect(anc.x0, anc.y1+10, anc.x0+560, anc.y1+380)
        txt  = text_in_rect(page, area)
        low  = txt.lower()

        # modelo por keywords
        modelo=""
        for kw in MODEL_KEYWORDS:
            if kw in low:
                modelo = "SynWay" if "syn" in kw else "aligera"
                break

        # status por choices
        status=""
        for ch in STATUS_CHOICES:
            if ch in low:
                status = ch
                break
        if status == "instalado pelo tecnico": status = "instalado pelo técnico"
        if status == "retirado pelo tecnico": status = "retirado pelo técnico"
        if status == "spare tecnico": status = "spare técnico"

        # serial: prioriza S/N, depois candidatos válidos (com dígito)
        numero_serie = ""
        m_sn = re.search(r"S/?N[:\s\-]*([A-Z0-9\-]{5,})", txt, flags=re.I)
        if m_sn:
            cand = clean_value(m_sn.group(1))
            if _looks_like_serial(cand):
                numero_serie = cand
        if not numero_serie:
            # pega todos tokens alfanum >=6 e filtra com dígito + não-stopword
            cands = re.findall(r"\b([A-Z0-9\-]{6,})\b", txt, flags=re.I)
            cands = [c for c in cands if _looks_like_serial(c)]
            if cands:
                numero_serie = cands[0]

        # tipo
        m_tipo = re.search(r"Tipo\s*:\s*([^\n|]+)", txt, flags=re.I)
        tipo = clean_value(m_tipo.group(1)) if m_tipo else ""

        if any([tipo,numero_serie,modelo,status]):
            out["tipo"]=tipo; out["numero_serie"]=numero_serie; out["modelo"]=modelo; out["status"]=status
            return out

    return out

def extract_equipamento_principal(doc) -> Dict[str,str]:
    """Procura o bloco de equipamentos em TODAS as páginas, retorna o primeiro válido."""
    for pno in range(doc.page_count):
        page = doc[pno]
        info = _extract_equipment_block_from_page(page)
        # retorna assim que tiver ao menos serial ou status/modelo
        if info.get("numero_serie") or info.get("modelo") or info.get("status") or info.get("tipo"):
            return info
    return {"tipo":"","numero_serie":"","modelo":"","status":""}

# -------- máscara --------
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
    obs = observacoes_raw or ""

    # Produtivo
    m_prod = re.search(r"\bProdutivo:\s*([^\n\r]+)", obs, flags=re.I)
    produtivo = clean_value(m_prod.group(1)) if m_prod else ""

    # Suporte
    m_sup = re.search(r"acompanhado pelo analista\s+([A-Za-zÀ-ÿ\s]+)", obs, flags=re.I)
    suporte_mam = clean_value(m_sup.group(1)) if m_sup else clean_value(suporte_mam_hint)

    # BA
    m_ba = re.search(r"\bBA:\s*([A-Za-z0-9\-_/]+)", acao or "", flags=re.I)
    ba_num = clean_value(m_ba.group(1)) if m_ba else ""

    # Descrição sem a linha de produtivo
    obs_clean = re.sub(r"(?im)^.*\bProdutivo:\s*[^\n\r]+$", "", obs).strip()

    tf = (teste_final or "").upper()
    tf_answer = "sim" if tf=="S" else ("não" if tf=="N" else "não")
    testado_com = ("Teste final realizado com o CPE do cliente conectado ao circuito, validação de camada 3 concluída."
                   if tf_answer=="sim" else
                   "Sem teste final com o equipamento do cliente no momento do atendimento.")

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

    lines=[]
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
    lines.append(f"DESCRIÇÃO: {obs_clean}")
    return "\n".join(lines)

# -------- API pública --------
def extract_from_pdf(path: str) -> str:
    doc = fitz.open(path)
    try:
        page1 = doc[0]
        numero = extract_numero_chamado(page1)
        ident  = extract_identificacao_pagina1(page1)

        equip  = extract_equipamento_principal(doc)

        page_obs = doc[1] if doc.page_count>=2 else doc[0]
        obs  = extract_observacoes(page_obs) or extract_observacoes(page1)
        prob = extract_problema(page_obs)    or extract_problema(page1)
        acao = extract_acao(page_obs)        or extract_acao(page1)

        mask = build_mask(
            numero_chamado=numero,
            equip=equip,
            tecnico=ident.get("tecnico",""),
            cliente_ciente=ident.get("cliente_ciente",""),
            contato=ident.get("contato",""),
            suporte_mam_hint="",
            teste_final=ident.get("teste_final",""),
            observacoes_raw=obs,
            problema=prob,
            acao=acao,
        )
        return mask
    finally:
        doc.close()
