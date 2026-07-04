import math
import re
import io
import os
import base64
import pandas as pd
from datetime import datetime, date
import traceback
import streamlit as st

# ==============================
# VERSÃO
# ==============================
VERSAO = "V1"

# ==============================
# TABELAS E FUNÇÕES AUXILIARES
# ==============================

TABELA_IR_TRADICIONAL = [
    (2428.80, 0.00, 0.00),
    (2826.65, 0.075, 182.16),
    (3751.05, 0.15, 394.16),
    (4664.68, 0.225, 675.49),
    (None,    0.275, 908.73),
]

TABELA_IR_ATE_042025 = [
    (2259.20, 0.00, 0.00),
    (2826.65, 0.075, 169.44),
    (3751.05, 0.15, 381.44),
    (4664.68, 0.225, 662.77),
    (None,    0.275, 896.00),
]

VALOR_DEP = 189.59
DATA_CORTE_TABELA_IR = date(2025, 5, 1)
DEDUCAO_SIMPLIFICADA_2026 = 607.20
TETO_INSS_2025 = 8157.41
TETO_INSS_2026 = 8475.55


def excel_date_to_datetime(value):
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass
    if isinstance(value, datetime):
        return value.replace(tzinfo=None)
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            ts = pd.to_datetime(value, unit="D", origin="1899-12-30", errors="raise")
            return ts.to_pydatetime().replace(tzinfo=None)
        except Exception:
            return None
    if isinstance(value, str):
        s = value.strip()
        if not s:
            return None
        dt = pd.to_datetime(s, errors="coerce", dayfirst=True)
        if pd.isna(dt):
            dt = pd.to_datetime(s, errors="coerce", dayfirst=False)
        if pd.isna(dt):
            return None
        return dt.to_pydatetime().replace(tzinfo=None)
    dt = pd.to_datetime(value, errors="coerce", dayfirst=True)
    if pd.isna(dt):
        return None
    return dt.to_pydatetime().replace(tzinfo=None)


def truncar(valor, casas=2):
    if valor is None:
        return 0.0
    fator = 10 ** casas
    return math.floor(float(valor) * fator) / fator


def limpar_negativo(valor):
    if valor is None:
        return 0.0
    try:
        v = float(valor)
    except Exception:
        return 0.0
    return 0.0 if v < 0 else v


def fmt_num(valor, tamanho, casas=2, permitir_negativo=False):
    if valor is None:
        valor = 0.0
    try:
        if pd.isna(valor):
            valor = 0.0
    except Exception:
        pass
    valor = truncar(valor, casas=casas)
    if not permitir_negativo and valor < 0:
        valor = 0.0
    inteiro = int(valor * (10 ** casas))
    s = f"{inteiro:d}"
    s = s[-tamanho:] if len(s) > tamanho else s.zfill(tamanho)
    return s


def fmt_int(valor, tamanho):
    if valor is None:
        valor = 0
    try:
        if pd.isna(valor):
            valor = 0
    except Exception:
        pass
    inteiro = int(valor)
    s = f"{inteiro:d}"
    s = s[-tamanho:] if len(s) > tamanho else s.zfill(tamanho)
    return s


def fmt_str(texto, tamanho):
    if texto is None:
        texto = ""
    try:
        if pd.isna(texto):
            texto = ""
    except Exception:
        pass
    return str(texto).ljust(tamanho)[:tamanho]


def competencia_aaaamm(data_excel):
    dt = excel_date_to_datetime(data_excel)
    if dt is None:
        return "000000"
    return dt.strftime("%Y%m")


def ultimo_dia_competencia(data_excel):
    dt = excel_date_to_datetime(data_excel)
    if dt is None:
        return None
    ano, mes = dt.year, dt.month
    prox = datetime(ano + 1, 1, 1) if mes == 12 else datetime(ano, mes + 1, 1)
    return prox - pd.Timedelta(days=1)


def tabela_ir_por_data_pagto(data_pagto_dt):
    if data_pagto_dt is None:
        return TABELA_IR_TRADICIONAL
    return TABELA_IR_ATE_042025 if data_pagto_dt.date() < DATA_CORTE_TABELA_IR else TABELA_IR_TRADICIONAL


def deducao_simplificada_por_data_pagto(data_pagto_dt):
    if data_pagto_dt is None:
        return 0.0
    return 564.80 if data_pagto_dt.date() < DATA_CORTE_TABELA_IR else 607.20


def deducao_simplificada_por_data_pagto_ou_ano(data_pagto_dt):
    if data_pagto_dt is None:
        return 0.0
    if data_pagto_dt.year >= 2026:
        return DEDUCAO_SIMPLIFICADA_2026
    return deducao_simplificada_por_data_pagto(data_pagto_dt)


def teto_inss_por_data_pagto(data_pagto_dt):
    if data_pagto_dt is None:
        return TETO_INSS_2026
    return TETO_INSS_2026 if data_pagto_dt.year >= 2026 else TETO_INSS_2025


def chave_acumulacao_mes(meta, reg, data_pagto_dt):
    competencia = (
        data_pagto_dt.strftime("%Y%m")
        if data_pagto_dt is not None
        else competencia_aaaamm(meta["competencia"])
    )
    return (int(meta["codigo_empresa"]), str(reg["cod_contrib"]).strip(), competencia)


def obter_rendimento_tributavel_irrf(bruto, esocial_int):
    bruto = limpar_negativo(bruto)
    if bruto <= 0:
        return 0.0
    if esocial_int in (711, 731, 734):
        return truncar(bruto * 0.60, casas=2)
    if esocial_int == 712:
        return truncar(bruto * 0.10, casas=2)
    return truncar(bruto, casas=2)


def calcular_irrf_tabela(base, tabela):
    if base is None or base <= 0:
        return 0.0
    aliquota, deducao = 0.0, 0.0
    for limite, aliq, ded in tabela:
        if limite is None or base <= limite:
            aliquota, deducao = aliq, ded
            break
    irrf = truncar(truncar(base * aliquota, casas=2) - deducao, casas=2)
    return max(irrf, 0.0)


def reducao_mensal_2026(rendimento_tributavel):
    if rendimento_tributavel is None:
        return 0.0
    try:
        rt = float(rendimento_tributavel)
    except Exception:
        return 0.0
    if rt <= 0:
        return 0.0
    if rt <= 5000.00:
        return 312.89
    if rt <= 7350.00:
        return truncar(978.62 - truncar(0.133145 * rt, casas=2), casas=2)
    return 0.0


def calcular_irrf_2026_por_base(BC, rendimento_tributavel):
    if BC is None or BC <= 0:
        return 0.0
    ir_tabela = calcular_irrf_tabela(BC, TABELA_IR_TRADICIONAL)
    if ir_tabela <= 0:
        return 0.0
    red = reducao_mensal_2026(rendimento_tributavel)
    return max(truncar(ir_tabela - min(red, ir_tabela), casas=2), 0.0)


def calcular_irrf_acumulado_generico(
    rendimento_tributavel_acum, inss_dedutivel_acum,
    dependentes, ano_ir, tabela_ir, ded_simpl
):
    if rendimento_tributavel_acum is None or rendimento_tributavel_acum <= 0:
        return 0.0, 0.0
    dep_int = max(0, 0 if (dependentes is None or pd.isna(dependentes)) else int(dependentes))
    red_dep = truncar(dep_int * VALOR_DEP, casas=2)
    base_legal = max(truncar(rendimento_tributavel_acum - inss_dedutivel_acum - red_dep, casas=2), 0.0)
    base_simpl = max(truncar(rendimento_tributavel_acum - ded_simpl, casas=2), 0.0)
    if ano_ir == 2026:
        ir_legal = calcular_irrf_2026_por_base(base_legal, rendimento_tributavel_acum)
        ir_simpl  = calcular_irrf_2026_por_base(base_simpl, rendimento_tributavel_acum)
    else:
        ir_legal = calcular_irrf_tabela(base_legal, tabela_ir)
        ir_simpl  = calcular_irrf_tabela(base_simpl, tabela_ir)
    if (ir_simpl < ir_legal) or (ir_simpl == ir_legal and base_simpl <= base_legal):
        return ir_simpl, base_simpl
    return ir_legal, base_legal


# ==============================
# LEITURA DO EXCEL (RPA)
# ==============================
def ler_planilha_rpa(caminho_excel, log):
    try:
        df = pd.read_excel(caminho_excel, sheet_name=0, header=None)
    except Exception as e:
        log.append(f"ERRO ao ler Excel: {e}")
        raise

    codigo_empresa = razao_social = cnpj = competencia = None

    for i in range(len(df)):
        c0 = df.iloc[i, 0]
        if pd.isna(c0):
            continue
        c0_str = str(c0).strip()
        prefixo = "RELAÇÃO DE RENDIMENTOS - RPA:"
        resto = c0_str[len(prefixo):].strip() if c0_str.startswith(prefixo) else c0_str
        if resto.startswith("Empresa"):
            codigo_empresa = df.iloc[i, 1]
        elif resto.startswith("Razão Social"):
            razao_social = df.iloc[i, 1]
        elif resto.startswith("CNPJ"):
            cnpj = df.iloc[i, 1]
        elif resto.startswith("Competencia"):
            competencia = df.iloc[i, 1]

    for campo, nome in [
        (codigo_empresa, "Codigo Empresa"),
        (razao_social,   "Razão Social"),
        (cnpj,           "CNPJ"),
        (competencia,    "Competencia"),
    ]:
        if campo is None or (isinstance(campo, float) and pd.isna(campo)):
            log.append(f"ERRO: '{nome}' não encontrado.")
            return None

    codigo_empresa = int(codigo_empresa)
    inicio = None
    tem_cpf = False
    ncol = df.shape[1]

    for i in range(len(df)):
        def cell(r, c):
            return None if c >= ncol else df.iloc[r, c]

        def cs(r, c):
            v = cell(r, c)
            return "" if (v is None or pd.isna(v)) else str(v).replace("RELAÇÃO DE RENDIMENTOS - RPA:", "").strip()

        if (cs(i,0)=="Código" and cs(i,1)=="Nome" and cs(i,2)=="CPF" and
            cs(i,3)=="Quantidade" and cs(i,4)=="Categoria" and cs(i,5)=="Próxima" and
            cs(i,6)=="Descrição" and cs(i,7)=="Rendimento" and cs(i,13)=="Data ISS"):
            inicio, tem_cpf = i + 2, True
            break
        if (cs(i,0)=="Código" and cs(i,1)=="Nome" and
            cs(i,2)=="Quantidade" and cs(i,3)=="Categoria" and
            cs(i,4)=="Próxima" and cs(i,5)=="Descrição" and cs(i,6)=="Rendimento"):
            inicio = i + 2
            break

    if inicio is None:
        log.append("ERRO: cabeçalho de contribuintes não encontrado.")
        return None

    def _num_or_zero(v):
        if v is None:
            return 0.0
        try:
            if pd.isna(v):
                return 0.0
        except Exception:
            pass
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return float(v)
        s = re.sub(r"[^0-9,\.\-]", "", str(v).strip())
        if s in ("", "-", ",", ".", "-.", "-,"):
            return 0.0
        if "." in s and "," in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")
        elif s.count(".") > 1:
            parts = s.split(".")
            s = "".join(parts[:-1]) + "." + parts[-1]
        try:
            return float(s)
        except Exception:
            return 0.0

    registros = []
    for i in range(inicio, len(df)):
        linha = df.iloc[i]
        cod_contrib = linha[0] if len(linha) > 0 else None
        if cod_contrib is None or pd.isna(cod_contrib):
            continue
        try:
            if tem_cpf:
                nome=linha[1]; dependentes=linha[3]; esocial=linha[4]
                rpa_num=linha[5]; atividade=linha[6]; bruto=linha[7]
                data_pagto=linha[8]; pensao=linha[9]; outros_desc=linha[10]
                outros_prov=linha[11]; perc_iss=linha[12]; data_iss=linha[13]
            else:
                nome=linha[1]; dependentes=linha[2]; esocial=linha[3]
                rpa_num=linha[4]; atividade=linha[5]; bruto=linha[6]
                data_pagto=linha[7]; pensao=linha[8]; outros_desc=linha[9]
                outros_prov=linha[10]; perc_iss=linha[11]; data_iss=linha[12]

            if bruto is None or pd.isna(bruto):
                log.append(f"Aviso: linha {i+1} sem BRUTO. Código: {cod_contrib}. Pulando.")
                continue

            registros.append({
                "cod_contrib": cod_contrib, "nome": nome,
                "dependentes": dependentes, "esocial": esocial,
                "rpa_num": rpa_num, "atividade": atividade,
                "bruto": _num_or_zero(bruto), "data_pagto": data_pagto,
                "pensao_alim": _num_or_zero(pensao),
                "outros_desc": _num_or_zero(outros_desc),
                "outros_prov": _num_or_zero(outros_prov),
                "perc_iss": _num_or_zero(perc_iss),
                "valor_iss": 0.0, "data_iss": data_iss, "linha_excel": i + 1,
            })
        except Exception as e:
            log.append(f"ERRO ao ler linha {i+1}: {e}")

    return {
        "codigo_empresa": codigo_empresa, "razao_social": razao_social,
        "cnpj": str(cnpj), "competencia": competencia, "registros": registros,
    }


# ==============================
# MONTAGEM DO REGISTRO TXT (266)
# ==============================
def montar_registro_lancamento(meta, reg, log, acum_mes):
    codigo_empresa   = meta["codigo_empresa"]
    competencia_data = meta["competencia"]
    competencia_str  = competencia_aaaamm(competencia_data)

    data_pagto_dt = excel_date_to_datetime(reg.get("data_pagto")) or ultimo_dia_competencia(competencia_data)
    data_pagto_str = "00000000" if data_pagto_dt is None else data_pagto_dt.strftime("%Y%m%d")
    ano_ir   = data_pagto_dt.year if data_pagto_dt is not None else None
    tabela_ir = tabela_ir_por_data_pagto(data_pagto_dt)
    ded_simpl = deducao_simplificada_por_data_pagto_ou_ano(data_pagto_dt)

    cod_contrib  = reg["cod_contrib"]
    dependentes  = reg["dependentes"]
    rpa_num      = reg["rpa_num"]
    atividade    = reg["atividade"]
    bruto        = limpar_negativo(reg["bruto"])
    perc_iss     = limpar_negativo(reg.get("perc_iss", 0.0))
    pensao_alim  = limpar_negativo(reg.get("pensao_alim", 0.0))
    outros_desc  = limpar_negativo(reg.get("outros_desc", 0.0))
    outros_prov  = limpar_negativo(reg.get("outros_prov", 0.0))

    dt_iss = excel_date_to_datetime(reg.get("data_iss"))
    data_venc_iss = "00000000" if dt_iss is None else dt_iss.strftime("%Y%m%d")

    esocial = reg.get("esocial")
    try:
        esocial_int = int(esocial) if not pd.isna(esocial) else None
    except Exception:
        esocial_int = None

    chave = chave_acumulacao_mes(meta, reg, data_pagto_dt)
    if chave not in acum_mes:
        acum_mes[chave] = {
            "base_inss_empresa": 0.0, "inss_retido_empresa": 0.0,
            "outras_fontes_base": 0.0, "rend_trib_irrf": 0.0,
            "inss_dedutivel_irrf": 0.0, "irrf_retido": 0.0, "dependentes": 0,
        }
    ac = acum_mes[chave]

    inss_frete_sest = inss_frete_senat = 0.0
    base_inss_registro_original = bruto
    aliquota_inss = 0.11

    if esocial_int in (712, 734):
        base_inss_registro_original = truncar(bruto * 0.20, casas=2)
        aliquota_inss = 0.20 if esocial_int == 734 else 0.11
        inss_frete_sest  = truncar(base_inss_registro_original * 0.015, casas=2)
        inss_frete_senat = truncar(base_inss_registro_original * 0.010, casas=2)

    teto_inss = teto_inss_por_data_pagto(data_pagto_dt)
    saldo_teto = max(truncar(teto_inss - max(truncar(ac.get("outras_fontes_base", 0.0), casas=2), 0.0), casas=2), 0.0)

    base_empresa_anterior = truncar(ac["base_inss_empresa"], casas=2)
    base_empresa_nova     = truncar(base_empresa_anterior + base_inss_registro_original, casas=2)
    base_inss_registro_limitada = max(
        truncar(min(base_empresa_nova, saldo_teto) - min(base_empresa_anterior, saldo_teto), casas=2), 0.0
    )

    inss = max(truncar(base_inss_registro_limitada * aliquota_inss, casas=2), 0.0)
    ac["base_inss_empresa"]    = base_empresa_nova
    ac["inss_retido_empresa"]  = truncar(ac["inss_retido_empresa"] + inss, casas=2)
    base_inss = base_inss_registro_limitada

    rendimento_tributavel_registro = obter_rendimento_tributavel_irrf(bruto, esocial_int)
    dep_out = max(0, 0 if (dependentes is None or pd.isna(dependentes)) else int(dependentes))
    deduz_inss = esocial_int in (711, 712)

    ac["rend_trib_irrf"]      = truncar(ac["rend_trib_irrf"] + rendimento_tributavel_registro, casas=2)
    ac["inss_dedutivel_irrf"] = truncar(ac["inss_dedutivel_irrf"] + inss, casas=2)
    ac["dependentes"]         = max(ac["dependentes"], dep_out)

    inss_dedutivel_acum = ac["inss_dedutivel_irrf"] if deduz_inss else 0.0

    if ano_ir in (2025, 2026):
        ir_total_mes, base_irrf_mes = calcular_irrf_acumulado_generico(
            ac["rend_trib_irrf"], inss_dedutivel_acum,
            ac["dependentes"], ano_ir, tabela_ir, ded_simpl
        )
    else:
        log.append(f"Aviso: ano desconhecido ({ano_ir}) para contrib {cod_contrib}; usando 2025.")
        ir_total_mes, base_irrf_mes = calcular_irrf_acumulado_generico(
            ac["rend_trib_irrf"], inss_dedutivel_acum,
            ac["dependentes"], 2025, tabela_ir, ded_simpl
        )

    ir_calculado = max(truncar(ir_total_mes - truncar(ac["irrf_retido"], casas=2), casas=2), 0.0)
    ac["irrf_retido"] = truncar(ac["irrf_retido"] + ir_calculado, casas=2)
    base_irrf = base_irrf_mes

    if perc_iss and float(perc_iss) != 0.0:
        valor_iss = truncar(bruto * (perc_iss / 100.0), casas=2)
    else:
        perc_iss, valor_iss = 0.0, 0.0

    valor_iss        = limpar_negativo(valor_iss)
    base_inss        = limpar_negativo(base_inss)
    inss_frete_sest  = limpar_negativo(inss_frete_sest)
    inss_frete_senat = limpar_negativo(inss_frete_senat)
    inss             = limpar_negativo(inss)
    base_irrf        = limpar_negativo(base_irrf)
    ir_calculado     = limpar_negativo(ir_calculado)

    try:
        registro = (
            fmt_int(codigo_empresa, 7)        +
            fmt_int(cod_contrib, 10)          +
            competencia_str                   +
            fmt_str(atividade, 100)           +
            fmt_int(rpa_num, 10)              +
            fmt_num(bruto, 11)                +
            fmt_num(perc_iss, 5)              +
            fmt_num(valor_iss, 11)            +
            data_venc_iss                     +
            fmt_num(base_inss, 11)            +
            fmt_num(inss_frete_sest, 8)       +
            fmt_num(inss_frete_senat, 8)      +
            fmt_num(inss, 8)                  +
            fmt_num(pensao_alim, 11)          +
            fmt_num(outros_desc, 11)          +
            fmt_num(outros_prov, 11)          +
            data_pagto_str                    +
            fmt_num(base_irrf, 11)            +
            fmt_int(dep_out, 3)               +
            fmt_num(ir_calculado, 8)
        )
    except Exception as e:
        log.append(f"ERRO ao montar registro do contrib {cod_contrib}: {e}")
        return None

    if len(registro) != 266:
        log.append(f"ERRO: Registro com tamanho {len(registro)} (esperado 266). Cód={codigo_empresa}, contrib={cod_contrib}")
        return None

    return registro


# ==============================
# GERAÇÃO DO TXT
# ==============================
def gerar_txt_streamlit(arquivo_bytes, log):
    try:
        meta = ler_planilha_rpa(io.BytesIO(arquivo_bytes), log)
        if meta is None:
            log.append("ERRO: Nenhum metadado/registro válido. Abortando.")
            return None

        meta["registros"].sort(key=lambda r: (
            str(r.get("cod_contrib", "")),
            excel_date_to_datetime(r.get("data_pagto")) or ultimo_dia_competencia(meta["competencia"]),
            int(r.get("rpa_num") or 0),
            int(r.get("linha_excel") or 0),
        ))

        linhas_txt = []
        acum_mes   = {}
        for reg in meta["registros"]:
            linha = montar_registro_lancamento(meta, reg, log, acum_mes)
            if linha is not None:
                linhas_txt.append(linha)

        if any(str(l).startswith("ERRO") for l in log):
            log.append("ERRO: Geração cancelada. TXT NÃO foi gerado.")
            return None
        if not linhas_txt:
            log.append("ERRO: Nenhum registro válido foi gerado para o TXT.")
            return None

        log.append(f"Arquivo TXT gerado com {len(linhas_txt)} registros.")
        return linhas_txt

    except Exception:
        log.append("ERRO FATAL durante a geração do arquivo.")
        log.append(traceback.format_exc())
        return None


# ==============================
# UTILITÁRIO — lê bgr_base64.txt
# ==============================
def carregar_bgr_bytes():
    caminho = os.path.join(os.path.dirname(os.path.abspath(__file__)), "bgr_base64.txt")
    if not os.path.exists(caminho):
        return None, "Arquivo bgr_base64.txt não encontrado na pasta do app."
    try:
        with open(caminho, "r", encoding="utf-8") as f:
            conteudo = f.read().strip()
        if "," in conteudo and conteudo.startswith("data:"):
            conteudo = conteudo.split(",", 1)[1]
        return base64.b64decode(conteudo), None
    except Exception as e:
        return None, f"Erro ao decodificar bgr_base64.txt: {e}"


# ==============================
# CSS GLOBAL — Tema TR / Domínio Escuro
# ==============================
TR_DARK_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

/* ── Reset e base ───────────────────────────────────── */
*, *::before, *::after { box-sizing: border-box; }

[data-testid="stAppViewContainer"],
[data-testid="stMain"], .main {
    background-color: #0D1117 !important;
}
[data-testid="stSidebar"] {
    background-color: #0D1117 !important;
    border-right: 1px solid #21262D !important;
}
[data-testid="stSidebarContent"] {
    padding: 0 !important;
}

html, body, [class*="css"] {
    font-family: 'Inter', 'Segoe UI', sans-serif !important;
    color: #E6EDF3 !important;
}
h1,h2,h3,h4,h5,h6 {
    color: #E6EDF3 !important;
    font-weight: 700 !important;
}

/* ── Botão primário ─────────────────────────────────── */
[data-testid="stButton"] button[kind="primary"] {
    background: linear-gradient(135deg, #FF6200, #E05500) !important;
    color: #fff !important;
    border: none !important;
    border-radius: 6px !important;
    font-weight: 700 !important;
    font-size: 14px !important;
    box-shadow: 0 2px 10px rgba(255,98,0,.4) !important;
    transition: opacity .2s, transform .1s !important;
}
[data-testid="stButton"] button[kind="primary"]:hover  { opacity:.88 !important; transform:translateY(-1px) !important; }
[data-testid="stButton"] button[kind="primary"]:active { transform:translateY(0) !important; }
[data-testid="stButton"] button[kind="primary"]:disabled {
    background: #2D333B !important; color: #484F58 !important; box-shadow:none !important;
}

/* ── Botão secundário ───────────────────────────────── */
[data-testid="stButton"] button[kind="secondary"] {
    background-color: #21262D !important;
    color: #E6EDF3 !important;
    border: 1px solid #30363D !important;
    border-radius: 6px !important;
    font-weight: 500 !important;
    transition: background-color .2s !important;
}
[data-testid="stButton"] button[kind="secondary"]:hover {
    background-color: #2D333B !important; border-color: #484F58 !important;
}

/* ── Download button ────────────────────────────────── */
[data-testid="stDownloadButton"] button {
    background: linear-gradient(135deg, #006AFF, #0054CC) !important;
    color: #fff !important;
    border: none !important;
    border-radius: 6px !important;
    font-weight: 700 !important;
    box-shadow: 0 2px 10px rgba(0,106,255,.4) !important;
    transition: opacity .2s !important;
}
[data-testid="stDownloadButton"] button:hover { opacity:.88 !important; }

/* ── File uploader ──────────────────────────────────── */
[data-testid="stFileUploader"] {
    background-color: #161B22 !important;
    border: 1.5px dashed #30363D !important;
    border-radius: 8px !important;
    transition: border-color .2s !important;
}
[data-testid="stFileUploader"]:hover { border-color: #FF6200 !important; }

/* ── Divisor nativo ─────────────────────────────────── */
hr { border-color: #21262D !important; }

/* ── Scrollbar ──────────────────────────────────────── */
::-webkit-scrollbar { width:6px; height:6px; }
::-webkit-scrollbar-track { background:#161B22; }
::-webkit-scrollbar-thumb { background:#30363D; border-radius:3px; }
::-webkit-scrollbar-thumb:hover { background:#484F58; }
</style>
"""


# ==============================
# SIDEBAR
# ==============================
def render_sidebar():
    with st.sidebar:

        # ── Cabeçalho da sidebar ─────────────────────────
        st.markdown("""
            <div style="
                background: linear-gradient(180deg,#1C2128 0%,#161B22 100%);
                border-bottom: 1px solid #21262D;
                padding: 20px 16px 16px;
                margin-bottom: 4px;
            ">
                <div style="
                    font-size:9px; font-weight:800; letter-spacing:.18em;
                    color:#FF6200; text-transform:uppercase; margin-bottom:4px;
                ">Thomson Reuters</div>
                <div style="font-size:16px; font-weight:700; color:#E6EDF3; line-height:1.2;">
                    Domínio Sistemas
                </div>
                <div style="font-size:11px; color:#8B949E; margin-top:3px;">
                    Gerador RPA &nbsp;·&nbsp; Utilitários
                </div>
            </div>
        """, unsafe_allow_html=True)

        # ── Seção BGR ────────────────────────────────────
        st.markdown("""
            <div style="padding: 16px 16px 0;">
                <div style="
                    display:flex; align-items:center; gap:8px;
                    border-left:3px solid #006AFF;
                    padding-left:10px; margin-bottom:12px;
                ">
                    <span style="font-size:18px;">🖼️</span>
                    <div>
                        <div style="font-size:13px; font-weight:700; color:#E6EDF3;">
                            Arquivo BGR
                        </div>
                        <div style="font-size:11px; color:#8B949E;">
                            Background · Domínio Sistemas
                        </div>
                    </div>
                </div>
            </div>
        """, unsafe_allow_html=True)

        # Descrição em texto nativo (sem HTML aninhado)
        st.caption(
            "Instale o plano de fundo personalizado no Domínio Sistemas via "
            "**Utilitários → Personalizar → BGR**."
        )

        # Botão de download BGR
        bgr_bytes, bgr_erro = carregar_bgr_bytes()

        with st.container():
            if bgr_bytes is not None:
                st.download_button(
                    label="⬇  Baixar bgr_base64.txt",
                    data=bgr_bytes,
                    file_name="bgr_base64.txt",
                    mime="text/plain",
                    use_container_width=True,
                    help="Clique para baixar o arquivo de background do Domínio Sistemas",
                )
                st.markdown(
                    f"<div style='font-size:11px;color:#3FB950;margin-top:4px;'>"
                    f"✔ Arquivo disponível · {len(bgr_bytes):,} bytes</div>",
                    unsafe_allow_html=True,
                )
            else:
                st.markdown(
                    f"<div style='background:#1A0D0D;border:1px solid #DA3633;"
                    f"border-radius:6px;padding:10px 12px;font-size:12px;"
                    f"color:#F85149;'>"
                    f"⚠ BGR não encontrado<br>"
                    f"<span style='color:#8B949E;font-size:11px;'>{bgr_erro}</span>"
                    f"</div>",
                    unsafe_allow_html=True,
                )

        st.markdown("<div style='margin:16px 0 4px;'>", unsafe_allow_html=True)
        st.divider()
        st.markdown("</div>", unsafe_allow_html=True)

        # ── Instruções de instalação (componentes nativos) ─
        st.markdown(
            "<div style='font-size:11px;font-weight:700;letter-spacing:.1em;"
            "color:#8B949E;text-transform:uppercase;padding:0 0 10px;'>"
            "📋 Como instalar o BGR</div>",
            unsafe_allow_html=True,
        )

        passos = [
            ("🟠", "Baixe o arquivo **bgr_base64.txt** acima"),
            ("🟠", "Abra o **Domínio Sistemas**"),
            ("🟠", "Acesse **Utilitários → Personalizar**"),
            ("🟠", "Importe o **.txt** na aba **BGR**"),
            ("🟢", "Reinicie o sistema para aplicar"),
        ]
        for icone, texto in passos:
            st.markdown(
                f"<div style='display:flex;align-items:flex-start;gap:8px;"
                f"margin-bottom:8px;'>"
                f"<span style='font-size:12px;margin-top:1px;'>{icone}</span>"
                f"<span style='font-size:12px;color:#C9D1D9;line-height:1.5;'>{texto}</span>"
                f"</div>",
                unsafe_allow_html=True,
            )

        # ── Rodapé da sidebar ─────────────────────────────
        st.divider()
        st.markdown(
            f"<div style='text-align:center;font-size:11px;color:#484F58;line-height:1.7;'>"
            f"Gerador RPA · <span style='color:#FF6200;font-weight:700;'>{VERSAO}</span><br>"
            f"Thomson Reuters · Domínio Sistemas"
            f"</div>",
            unsafe_allow_html=True,
        )


# ==============================
# MAIN
# ==============================
def main():
    st.set_page_config(
        page_title=f"Gerador TXT — RPA | {VERSAO}",
        page_icon="📄",
        layout="wide",
    )

    st.markdown(TR_DARK_CSS, unsafe_allow_html=True)
    render_sidebar()

    # ── Header ───────────────────────────────────────────
    st.markdown(f"""
        <div style="
            background:linear-gradient(135deg,#161B22 0%,#1C2128 100%);
            border:1px solid #21262D; border-top:3px solid #FF6200;
            border-radius:10px; padding:26px 28px 18px; margin-bottom:24px;
        ">
            <div style="display:flex;align-items:center;gap:14px;margin-bottom:4px;">
                <span style="font-size:26px;">📄</span>
                <div style="flex:1;">
                    <div style="font-size:10px;font-weight:800;letter-spacing:.14em;
                                color:#FF6200;text-transform:uppercase;margin-bottom:2px;">
                        Thomson Reuters · Domínio Sistemas
                    </div>
                    <div style="font-size:20px;font-weight:700;color:#E6EDF3;letter-spacing:-.02em;">
                        Gerador de Arquivo TXT — RPA
                    </div>
                </div>
                <div style="background:#21262D;border:1px solid #30363D;border-radius:20px;
                            padding:3px 14px;font-size:12px;font-weight:700;
                            color:#8B949E;letter-spacing:.06em;">
                    {VERSAO}
                </div>
            </div>
            <p style="margin:8px 0 0 40px;font-size:13px;color:#8B949E;line-height:1.5;">
                Importe o Excel de origem e clique em
                <strong style="color:#E6EDF3;">Gerar arquivo TXT</strong>
                para processar os lançamentos RPA.
            </p>
        </div>
    """, unsafe_allow_html=True)

    # ── Session state ─────────────────────────────────────
    if "log"        not in st.session_state: st.session_state.log        = [f"✔ Aplicação pronta — {VERSAO}"]
    if "txt_gerado" not in st.session_state: st.session_state.txt_gerado = None
    if "meta_info"  not in st.session_state: st.session_state.meta_info  = None

    # ── Upload ────────────────────────────────────────────
    st.markdown("""
        <div style="background:#161B22;border:1px solid #21262D;border-radius:8px;
                    padding:16px 16px 2px;margin-bottom:14px;">
            <div style="font-size:10px;font-weight:700;letter-spacing:.1em;
                        color:#8B949E;text-transform:uppercase;margin-bottom:8px;">
                📂 Arquivo de entrada
            </div>
    """, unsafe_allow_html=True)

    arquivo = st.file_uploader(
        "Excel de entrada",
        type=["xlsx", "xls"],
        help="Planilha de Relação de Rendimentos — RPA",
        label_visibility="collapsed",
    )
    st.markdown("</div>", unsafe_allow_html=True)

    # ── Botões ────────────────────────────────────────────
    c1, c2 = st.columns([4, 1])
    with c1:
        gerar  = st.button("▶  Gerar arquivo TXT", disabled=(arquivo is None),
                           use_container_width=True, type="primary")
    with c2:
        limpar = st.button("🗑  Limpar", use_container_width=True, type="secondary")

    if limpar:
        st.session_state.log        = ["🗑 Limpo. Pronto para nova operação."]
        st.session_state.txt_gerado = None
        st.session_state.meta_info  = None
        st.rerun()

    # ── Processamento ─────────────────────────────────────
    if gerar and arquivo is not None:
        st.session_state.log        = ["⏳ Iniciando processamento..."]
        st.session_state.txt_gerado = None
        st.session_state.meta_info  = None

        arquivo_bytes = arquivo.read()
        with st.spinner("Processando lançamentos RPA…"):
            linhas = gerar_txt_streamlit(arquivo_bytes, st.session_state.log)

        if linhas:
            conteudo = "\n".join(linhas) + "\n"
            st.session_state.txt_gerado = conteudo.encode("latin-1", errors="replace")
            try:
                _meta = ler_planilha_rpa(io.BytesIO(arquivo_bytes), [])
                if _meta:
                    st.session_state.meta_info = {
                        "empresa":    _meta.get("razao_social", "—"),
                        "cnpj":       _meta.get("cnpj", "—"),
                        "competencia": competencia_aaaamm(_meta.get("competencia")),
                        "registros":  len(linhas),
                    }
            except Exception:
                pass
        st.rerun()

    # ── Cards de resultado ────────────────────────────────
    if st.session_state.txt_gerado is not None:
        if st.session_state.meta_info:
            m = st.session_state.meta_info
            comp = (f"{m['competencia'][4:6]}/{m['competencia'][:4]}"
                    if len(m["competencia"]) == 6 else m["competencia"])
            c_emp, c_cnpj, c_comp, c_reg = st.columns(4)
            for col, label, valor, cor in [
                (c_emp,  "Empresa",     m["empresa"],   "#FF6200"),
                (c_cnpj, "CNPJ",        m["cnpj"],      "#006AFF"),
                (c_comp, "Competência", comp,           "#D29922"),
                (c_reg,  "Registros",   m["registros"], "#3FB950"),
            ]:
                with col:
                    st.markdown(f"""
                        <div style="background:#161B22;border:1px solid #21262D;
                                    border-top:2px solid {cor};border-radius:8px;
                                    padding:14px;margin-bottom:12px;">
                            <div style="font-size:10px;font-weight:700;letter-spacing:.1em;
                                        color:#8B949E;text-transform:uppercase;margin-bottom:4px;">
                                {label}
                            </div>
                            <div style="font-size:{'20px' if label=='Registros' else '13px'};
                                        font-weight:700;color:{cor if label=='Registros' else '#E6EDF3'};
                                        overflow:hidden;text-overflow:ellipsis;white-space:nowrap;"
                                 title="{valor}">{valor}</div>
                        </div>
                    """, unsafe_allow_html=True)

        st.markdown("""
            <div style="background:#0D2818;border:1px solid #238636;border-left:4px solid #3FB950;
                        border-radius:8px;padding:12px 18px;margin-bottom:12px;
                        display:flex;align-items:center;gap:10px;">
                <span style="font-size:18px;">✅</span>
                <span style="font-size:14px;font-weight:600;color:#3FB950;">
                    Arquivo gerado com sucesso! Clique abaixo para baixar.
                </span>
            </div>
        """, unsafe_allow_html=True)

        st.download_button(
            label="⬇  Baixar arquivo TXT",
            data=st.session_state.txt_gerado,
            file_name="saida_rpa.txt",
            mime="text/plain",
            use_container_width=True,
            type="primary",
        )

    # ── Log ───────────────────────────────────────────────
    st.markdown("<div style='margin-top:24px;'></div>", unsafe_allow_html=True)

    tem_erro  = any(str(l).startswith("ERRO")  for l in st.session_state.log)
    tem_aviso = any(str(l).startswith("Aviso") for l in st.session_state.log)

    if tem_erro:
        acc, brd, bg, ico, lbl = "#F85149","#DA3633","#1A0D0D","🔴","Erros detectados"
    elif tem_aviso:
        acc, brd, bg, ico, lbl = "#D29922","#9E6A03","#1A1500","🟡","Avisos"
    else:
        acc, brd, bg, ico, lbl = "#3FB950","#238636","#0D1A10","🟢","OK"

    st.markdown(f"""
        <div style="background:#161B22;border:1px solid #21262D;
                    border-top:2px solid {acc};border-radius:8px;overflow:hidden;">
            <div style="background:#1C2128;padding:10px 16px;border-bottom:1px solid #21262D;
                        display:flex;align-items:center;gap:8px;">
                <span>{ico}</span>
                <span style="font-size:11px;font-weight:700;letter-spacing:.08em;
                             color:#8B949E;text-transform:uppercase;">
                    Log de processamento — {lbl}
                </span>
            </div>
            <div style="background:{bg};border-left:3px solid {brd};margin:12px;
                        border-radius:4px;padding:12px 14px;
                        font-family:'Cascadia Code','Fira Code','Consolas',monospace;
                        font-size:12.5px;line-height:1.7;white-space:pre-wrap;
                        max-height:280px;overflow-y:auto;color:#C9D1D9;">
{chr(10).join(str(l) for l in st.session_state.log)}
            </div>
        </div>
    """, unsafe_allow_html=True)

    # ── Rodapé ────────────────────────────────────────────
    st.markdown(f"""
        <div style="margin-top:32px;padding-top:14px;border-top:1px solid #21262D;
                    display:flex;justify-content:space-between;align-items:center;">
            <span style="font-size:11px;color:#484F58;">
                Thomson Reuters · Domínio Sistemas · Gerador RPA
            </span>
            <span style="font-size:11px;font-weight:800;
                         letter-spacing:.08em;color:#FF6200;">{VERSAO}</span>
        </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
