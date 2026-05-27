import math
import re
import io
import pandas as pd
from datetime import datetime, date
import traceback
import streamlit as st

# ==============================
# VERSÃO
# ==============================
VERSAO = "V3.6"

# ==============================
# TEMA TR
# ==============================
def apply_tr_theme():
    st.markdown("""
        <style>
        html, body, [class*="css"] {
            font-family: 'Segoe UI', 'Arial', sans-serif;
            color: #444444;
        }
        h1, h2, h3 {
            color: #FF8000;
            font-weight: 700;
        }
        section[data-testid="stSidebar"] {
            background-color: #444444;
            color: #FFFFFF;
        }
        section[data-testid="stSidebar"] * {
            color: #FFFFFF !important;
        }
        .stButton > button {
            background-color: #FF8000;
            color: #FFFFFF;
            border: none;
            border-radius: 4px;
            font-weight: bold;
        }
        .stButton > button:hover {
            background-color: #D64001;
            color: #FFFFFF;
        }
        hr {
            border-color: #FF8000;
        }
        [data-testid="metric-container"] {
            background-color: #E9E9E9;
            border-left: 4px solid #FF8000;
            border-radius: 4px;
            padding: 10px;
        }
        </style>
    """, unsafe_allow_html=True)


# ==============================
# TABELAS E FUNÇÕES AUXILIARES
# ==============================

TABELA_IR_TRADICIONAL = [
    (2428.80, 0.00,   0.00),
    (2826.65, 0.075, 182.16),
    (3751.05, 0.15,  394.16),
    (4664.68, 0.225, 675.49),
    (None,    0.275, 908.73),
]

TABELA_IR_ATE_042025 = [
    (2259.20, 0.00,   0.00),
    (2826.65, 0.075, 169.44),
    (3751.05, 0.15,  381.44),
    (4664.68, 0.225, 662.77),
    (None,    0.275, 896.00),
]

VALOR_DEP                  = 189.59
DATA_CORTE_TABELA_IR       = date(2025, 5, 1)
DEDUCAO_SIMPLIFICADA_2026  = 607.20
TETO_INSS_2025             = 8157.41
TETO_INSS_2026             = 8475.55


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
    if data_pagto_dt.year >= 2026:
        return TETO_INSS_2026
    return TETO_INSS_2025


def chave_acumulacao_mes(meta, reg, data_pagto_dt):
    competencia = (
        data_pagto_dt.strftime("%Y%m")
        if data_pagto_dt is not None
        else competencia_aaaamm(meta["competencia"])
    )
    return (
        int(meta["codigo_empresa"]),
        str(reg["cod_contrib"]).strip(),
        competencia,
    )


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
    aliquota = deducao = 0.0
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


def calcular_irrf_mais_vantajoso_base100(base_bruta, dependentes, tabela, ded_simpl):
    if base_bruta is None or base_bruta <= 0:
        return 0.0, 0.0
    dep_int = 0 if (dependentes is None or pd.isna(dependentes)) else int(dependentes)
    red_dep = truncar(dep_int * VALOR_DEP, casas=2)
    base_geral = truncar(base_bruta - red_dep, casas=2)
    ir_geral   = calcular_irrf_tabela(base_geral, tabela)
    base_simpl = truncar(base_bruta - ded_simpl, casas=2)
    ir_simpl   = calcular_irrf_tabela(base_simpl, tabela)
    if (ir_simpl < ir_geral) or (ir_simpl == ir_geral and base_simpl <= base_geral):
        return ir_simpl, base_simpl
    return ir_geral, base_geral


def calcular_irrf_base60_legal(bruto, inss, dependentes, tabela):
    if bruto is None or bruto <= 0:
        return 0.0, 0.0
    base60  = truncar(bruto * 0.60, casas=2)
    dep_int = 0 if (dependentes is None or pd.isna(dependentes)) else int(dependentes)
    base    = truncar(base60 - inss - truncar(dep_int * VALOR_DEP, casas=2), casas=2)
    return calcular_irrf_tabela(base, tabela), base


def calcular_irrf_base60_mais_vantajoso_2025(bruto, inss, dependentes, tabela, ded_simpl):
    if bruto is None or bruto <= 0:
        return 0.0, 0.0
    base60     = truncar(bruto * 0.60, casas=2)
    ir_geral, base_geral = calcular_irrf_base60_legal(bruto, inss, dependentes, tabela)
    base_simpl = truncar(base60 - ded_simpl, casas=2)
    ir_simpl   = calcular_irrf_tabela(base_simpl, tabela)
    if (ir_simpl < ir_geral) or (ir_simpl == ir_geral and base_simpl <= base_geral):
        return ir_simpl, base_simpl
    return ir_geral, base_geral


def calcular_irrf_base10_legal(bruto, inss, dependentes, tabela):
    if bruto is None or bruto <= 0:
        return 0.0, 0.0
    base10  = truncar(bruto * 0.10, casas=2)
    dep_int = 0 if (dependentes is None or pd.isna(dependentes)) else int(dependentes)
    base    = truncar(base10 - inss - truncar(dep_int * VALOR_DEP, casas=2), casas=2)
    return calcular_irrf_tabela(base, tabela), base


def calcular_irrf_base10_mais_vantajoso_2025(bruto, inss, dependentes, tabela, ded_simpl):
    if bruto is None or bruto <= 0:
        return 0.0, 0.0
    base10     = truncar(bruto * 0.10, casas=2)
    ir_geral, base_geral = calcular_irrf_base10_legal(bruto, inss, dependentes, tabela)
    base_simpl = truncar(base10 - ded_simpl, casas=2)
    ir_simpl   = calcular_irrf_tabela(base_simpl, tabela)
    if (ir_simpl < ir_geral) or (ir_simpl == ir_geral and base_simpl <= base_geral):
        return ir_simpl, base_simpl
    return ir_geral, base_geral


def calcular_irrf_base60_mais_vantajoso_2026(bruto, inss, dependentes, ded_simpl, rendimento_tributavel):
    if bruto is None or bruto <= 0:
        return 0.0, 0.0
    base60     = truncar(bruto * 0.60, casas=2)
    dep_int    = 0 if (dependentes is None or pd.isna(dependentes)) else int(dependentes)
    base_legal = truncar(base60 - inss - truncar(dep_int * VALOR_DEP, casas=2), casas=2)
    ir_legal   = calcular_irrf_2026_por_base(base_legal, rendimento_tributavel)
    base_simpl = truncar(base60 - ded_simpl, casas=2)
    ir_simpl   = calcular_irrf_2026_por_base(base_simpl, rendimento_tributavel)
    if (ir_simpl < ir_legal) or (ir_simpl == ir_legal and base_simpl <= base_legal):
        return ir_simpl, base_simpl
    return ir_legal, base_legal


def calcular_irrf_base10_mais_vantajoso_2026(bruto, inss, dependentes, ded_simpl, rendimento_tributavel):
    if bruto is None or bruto <= 0:
        return 0.0, 0.0
    base10     = truncar(bruto * 0.10, casas=2)
    dep_int    = 0 if (dependentes is None or pd.isna(dependentes)) else int(dependentes)
    base_legal = truncar(base10 - inss - truncar(dep_int * VALOR_DEP, casas=2), casas=2)
    ir_legal   = calcular_irrf_2026_por_base(base_legal, rendimento_tributavel)
    base_simpl = truncar(base10 - ded_simpl, casas=2)
    ir_simpl   = calcular_irrf_2026_por_base(base_simpl, rendimento_tributavel)
    if (ir_simpl < ir_legal) or (ir_simpl == ir_legal and base_simpl <= base_legal):
        return ir_simpl, base_simpl
    return ir_legal, base_legal


def calcular_irrf_mais_vantajoso_2026_base100(base_bruta, dependentes, rendimento_tributavel, ded_simpl):
    if base_bruta is None or base_bruta <= 0:
        return 0.0, 0.0, "nenhum"
    dep_int    = 0 if (dependentes is None or pd.isna(dependentes)) else int(dependentes)
    base_legal = truncar(base_bruta - truncar(dep_int * VALOR_DEP, casas=2), casas=2)
    ir_legal   = calcular_irrf_2026_por_base(base_legal, rendimento_tributavel)
    base_simpl = truncar(base_bruta - ded_simpl, casas=2)
    ir_simpl   = calcular_irrf_2026_por_base(base_simpl, rendimento_tributavel)
    if (ir_simpl < ir_legal) or (ir_simpl == ir_legal and base_simpl <= base_legal):
        return ir_simpl, base_simpl, "simplificada"
    return ir_legal, base_legal, "legal"


def calcular_irrf_acumulado_generico(
    rendimento_tributavel_acum,
    inss_dedutivel_acum,
    dependentes,
    ano_ir,
    tabela_ir,
    ded_simpl,
):
    if rendimento_tributavel_acum is None or rendimento_tributavel_acum <= 0:
        return 0.0, 0.0
    dep_int    = max(0, 0 if (dependentes is None or pd.isna(dependentes)) else int(dependentes))
    red_dep    = truncar(dep_int * VALOR_DEP, casas=2)
    base_legal = max(truncar(rendimento_tributavel_acum - inss_dedutivel_acum - red_dep, casas=2), 0.0)
    base_simpl = max(truncar(rendimento_tributavel_acum - ded_simpl, casas=2), 0.0)
    if ano_ir == 2026:
        ir_legal = calcular_irrf_2026_por_base(base_legal, rendimento_tributavel_acum)
        ir_simpl = calcular_irrf_2026_por_base(base_simpl, rendimento_tributavel_acum)
    else:
        ir_legal = calcular_irrf_tabela(base_legal, tabela_ir)
        ir_simpl = calcular_irrf_tabela(base_simpl, tabela_ir)
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
        resto   = c0_str[len(prefixo):].strip() if c0_str.startswith(prefixo) else c0_str
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

    inicio  = None
    tem_cpf = False
    ncol    = df.shape[1]

    for i in range(len(df)):
        def cell(r, c):
            return None if c >= ncol else df.iloc[r, c]

        def cs(v):
            return "" if (v is None or pd.isna(v)) else str(v).replace("RELAÇÃO DE RENDIMENTOS - RPA:", "").strip()

        c0s  = cs(cell(i, 0))
        c1s  = cs(cell(i, 1))
        c2s  = cs(cell(i, 2))
        c3s  = cs(cell(i, 3))
        c4s  = cs(cell(i, 4))
        c5s  = cs(cell(i, 5))
        c6s  = cs(cell(i, 6))
        c7s  = cs(cell(i, 7))
        c13s = cs(cell(i, 13))

        if (
            c0s == "Código" and c1s == "Nome" and c2s == "CPF" and
            c3s == "Quantidade" and c4s == "Categoria" and c5s == "Próxima" and
            c6s == "Descrição" and c7s == "Rendimento" and c13s == "Data ISS"
        ):
            inicio  = i + 2
            tem_cpf = True
            break

        if (
            c0s == "Código" and c1s == "Nome" and
            c2s == "Quantidade" and c3s == "Categoria" and
            c4s == "Próxima" and c5s == "Descrição" and c6s == "Rendimento"
        ):
            inicio  = i + 2
            tem_cpf = False
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
        if s.count(".") > 1:
            parts = s.split(".")
            s = "".join(parts[:-1]) + "." + parts[-1]
        try:
            return float(s)
        except Exception:
            return 0.0

    registros = []

    for i in range(inicio, len(df)):
        linha       = df.iloc[i]
        cod_contrib = linha[0] if len(linha) > 0 else None
        if cod_contrib is None or pd.isna(cod_contrib):
            continue
        try:
            if tem_cpf:
                nome        = linha[1]
                dependentes = linha[3]
                esocial     = linha[4]
                rpa_num     = linha[5]
                atividade   = linha[6]
                bruto       = linha[7]
                data_pagto  = linha[8]
                pensao      = linha[9]
                outros_desc = linha[10]
                outros_prov = linha[11]
                perc_iss    = linha[12]
                data_iss    = linha[13]
            else:
                nome        = linha[1]
                dependentes = linha[2]
                esocial     = linha[3]
                rpa_num     = linha[4]
                atividade   = linha[5]
                bruto       = linha[6]
                data_pagto  = linha[7]
                pensao      = linha[8]
                outros_desc = linha[9]
                outros_prov = linha[10]
                perc_iss    = linha[11]
                data_iss    = linha[12]

            if bruto is None or pd.isna(bruto):
                log.append(f"Aviso: linha {i+1} sem BRUTO. Código: {cod_contrib}. Pulando.")
                continue

            registros.append({
                "cod_contrib": cod_contrib,
                "nome":        nome,
                "dependentes": dependentes,
                "esocial":     esocial,
                "rpa_num":     rpa_num,
                "atividade":   atividade,
                "bruto":       _num_or_zero(bruto),
                "data_pagto":  data_pagto,
                "pensao_alim": _num_or_zero(pensao),
                "outros_desc": _num_or_zero(outros_desc),
                "outros_prov": _num_or_zero(outros_prov),
                "perc_iss":    _num_or_zero(perc_iss),
                "valor_iss":   0.0,
                "data_iss":    data_iss,
                "linha_excel": i + 1,
            })
        except Exception as e:
            log.append(f"ERRO ao ler linha {i+1}: {e}")

    return {
        "codigo_empresa": codigo_empresa,
        "razao_social":   razao_social,
        "cnpj":           str(cnpj),
        "competencia":    competencia,
        "registros":      registros,
    }


# ==============================
# MONTAGEM DO REGISTRO TXT (266)
# ==============================
def montar_registro_lancamento(meta, reg, log, acum_mes):
    codigo_empresa   = meta["codigo_empresa"]
    competencia_data = meta["competencia"]
    competencia_str  = competencia_aaaamm(competencia_data)

    data_pagto_excel = reg.get("data_pagto")
    data_pagto_dt    = excel_date_to_datetime(data_pagto_excel)

    if data_pagto_dt is None:
        data_pagto_dt = ultimo_dia_competencia(competencia_data)

    data_pagto_str = "00000000" if data_pagto_dt is None else data_pagto_dt.strftime("%Y%m%d")
    ano_ir         = data_pagto_dt.year if data_pagto_dt is not None else None

    tabela_ir = tabela_ir_por_data_pagto(data_pagto_dt)
    ded_simpl = deducao_simplificada_por_data_pagto_ou_ano(data_pagto_dt)

    cod_contrib = reg["cod_contrib"]
    dependentes = reg["dependentes"]
    rpa_num     = reg["rpa_num"]
    atividade   = reg["atividade"]
    bruto       = limpar_negativo(reg["bruto"])

    perc_iss    = limpar_negativo(reg.get("perc_iss",    0.0))
    pensao_alim = limpar_negativo(reg.get("pensao_alim", 0.0))
    outros_desc = limpar_negativo(reg.get("outros_desc", 0.0))
    outros_prov = limpar_negativo(reg.get("outros_prov", 0.0))

    dt_iss        = excel_date_to_datetime(reg.get("data_iss"))
    data_venc_iss = "00000000" if dt_iss is None else dt_iss.strftime("%Y%m%d")

    esocial = reg.get("esocial")
    try:
        esocial_int = int(esocial) if not pd.isna(esocial) else None
    except Exception:
        esocial_int = None

    chave = chave_acumulacao_mes(meta, reg, data_pagto_dt)
    if chave not in acum_mes:
        acum_mes[chave] = {
            "base_inss_empresa":   0.0,
            "inss_retido_empresa": 0.0,
            "outras_fontes_base":  0.0,
            "rend_trib_irrf":      0.0,
            "inss_dedutivel_irrf": 0.0,
            "irrf_retido":         0.0,
            "dependentes":         0,
        }

    ac = acum_mes[chave]

    inss_frete_sest  = 0.0
    inss_frete_senat = 0.0

    # ------------------------------------------------------------------
    # BASE INSS
    # base_inss_registro_original → bruto integral (ou 20% para frete)
    # É o valor gravado no campo base_inss do TXT.
    # base_inss_registro_limitada → base após aplicação do teto acumulado
    # É o valor usado exclusivamente para calcular o valor do INSS.
    # ------------------------------------------------------------------
    base_inss_registro_original = bruto
    aliquota_inss = 0.11

    if esocial_int in (712, 734):
        base_inss_registro_original = truncar(bruto * 0.20, casas=2)
        aliquota_inss   = 0.20 if esocial_int == 734 else 0.11
        inss_frete_sest  = truncar(base_inss_registro_original * 0.015, casas=2)
        inss_frete_senat = truncar(base_inss_registro_original * 0.010, casas=2)

    teto_inss          = teto_inss_por_data_pagto(data_pagto_dt)
    outras_fontes_base = max(truncar(ac.get("outras_fontes_base", 0.0), casas=2), 0.0)
    saldo_teto         = max(truncar(teto_inss - outras_fontes_base, casas=2), 0.0)

    base_empresa_anterior = truncar(ac["base_inss_empresa"], casas=2)
    base_empresa_nova     = truncar(base_empresa_anterior + base_inss_registro_original, casas=2)

    base_limitada_anterior      = min(base_empresa_anterior, saldo_teto)
    base_limitada_nova          = min(base_empresa_nova,     saldo_teto)
    base_inss_registro_limitada = max(
        truncar(base_limitada_nova - base_limitada_anterior, casas=2), 0.0
    )

    inss = max(truncar(base_inss_registro_limitada * aliquota_inss, casas=2), 0.0)

    ac["base_inss_empresa"]   = base_empresa_nova
    ac["inss_retido_empresa"] = truncar(ac["inss_retido_empresa"] + inss, casas=2)

    # Campo base_inss no TXT → rendimento bruto original (sem limitação de teto)
    base_inss_saida = base_inss_registro_original

    # ------------------------------------------------------------------
    # IRRF acumulado
    # CORREÇÃO V3.6: INSS deduzido da base do IRRF para todos os eSociais
    # ------------------------------------------------------------------
    rendimento_tributavel_registro = obter_rendimento_tributavel_irrf(bruto, esocial_int)

    dep_out = max(0, 0 if (dependentes is None or pd.isna(dependentes)) else int(dependentes))

    ac["rend_trib_irrf"]      = truncar(ac["rend_trib_irrf"]      + rendimento_tributavel_registro, casas=2)
    ac["inss_dedutivel_irrf"] = truncar(ac["inss_dedutivel_irrf"] + inss,                           casas=2)
    ac["dependentes"]         = max(ac["dependentes"], dep_out)

    rendimento_tributavel_acum = ac["rend_trib_irrf"]
    inss_dedutivel_acum        = ac["inss_dedutivel_irrf"]   # ← CORREÇÃO: sempre deduz INSS
    dependentes_acum           = ac["dependentes"]

    _ano = ano_ir if ano_ir in (2025, 2026) else 2025
    if ano_ir not in (2025, 2026):
        log.append(
            f"Aviso: ano de pagamento desconhecido ({ano_ir}) para contrib "
            f"{cod_contrib}; usando regra 2025."
        )

    ir_total_mes, base_irrf_mes = calcular_irrf_acumulado_generico(
        rendimento_tributavel_acum=rendimento_tributavel_acum,
        inss_dedutivel_acum=inss_dedutivel_acum,
        dependentes=dependentes_acum,
        ano_ir=_ano,
        tabela_ir=tabela_ir,
        ded_simpl=ded_simpl,
    )

    irrf_ja_retido = truncar(ac["irrf_retido"], casas=2)
    ir_calculado   = max(truncar(ir_total_mes - irrf_ja_retido, casas=2), 0.0)
    ac["irrf_retido"] = truncar(ac["irrf_retido"] + ir_calculado, casas=2)

    base_irrf = base_irrf_mes

    # ISS
    if perc_iss and float(perc_iss) != 0.0:
        valor_iss = truncar(bruto * (perc_iss / 100.0), casas=2)
    else:
        perc_iss  = 0.0
        valor_iss = 0.0

    # Garantir não-negativos na saída
    valor_iss        = limpar_negativo(valor_iss)
    base_inss_saida  = limpar_negativo(base_inss_saida)
    inss_frete_sest  = limpar_negativo(inss_frete_sest)
    inss_frete_senat = limpar_negativo(inss_frete_senat)
    inss             = limpar_negativo(inss)
    base_irrf        = limpar_negativo(base_irrf)
    ir_calculado     = limpar_negativo(ir_calculado)

    # ------------------------------------------------------------------
    # Montagem dos campos posicionais (total = 266 caracteres)
    # ------------------------------------------------------------------
    try:
        campo_codigo_empresa   = fmt_int(codigo_empresa,  7)
        campo_codigo_contrib   = fmt_int(cod_contrib,    10)
        campo_competencia      = competencia_str
        campo_desc_atividade   = fmt_str(atividade,     100)
        campo_num_rpa          = fmt_int(rpa_num,        10)
        campo_rendimento_bruto = fmt_num(bruto,          11, casas=2, permitir_negativo=False)
        campo_percentual_iss   = fmt_num(perc_iss,        5, casas=2, permitir_negativo=False)
        campo_valor_iss        = fmt_num(valor_iss,      11, casas=2, permitir_negativo=False)
        campo_data_venc_iss    = data_venc_iss
        campo_base_inss        = fmt_num(base_inss_saida, 11, casas=2, permitir_negativo=False)
        campo_inss_frete_sest  = fmt_num(inss_frete_sest,  8, casas=2, permitir_negativo=False)
        campo_inss_frete_senat = fmt_num(inss_frete_senat, 8, casas=2, permitir_negativo=False)
        campo_valor_inss       = fmt_num(inss,             8, casas=2, permitir_negativo=False)
        campo_pensao_alim      = fmt_num(pensao_alim,     11, casas=2, permitir_negativo=False)
        campo_outros_desc      = fmt_num(outros_desc,     11, casas=2, permitir_negativo=False)
        campo_outros_prov      = fmt_num(outros_prov,     11, casas=2, permitir_negativo=False)
        campo_data_pagto       = data_pagto_str
        campo_base_irrf        = fmt_num(base_irrf,       11, casas=2, permitir_negativo=False)
        campo_qtd_dep_ir       = fmt_int(dep_out,          3)
        campo_valor_ir         = fmt_num(ir_calculado,     8, casas=2, permitir_negativo=False)

        registro = (
            campo_codigo_empresa   +   #   7
            campo_codigo_contrib   +   #  10
            campo_competencia      +   #   6
            campo_desc_atividade   +   # 100
            campo_num_rpa          +   #  10
            campo_rendimento_bruto +   #  11
            campo_percentual_iss   +   #   5
            campo_valor_iss        +   #  11
            campo_data_venc_iss    +   #   8
            campo_base_inss        +   #  11
            campo_inss_frete_sest  +   #   8
            campo_inss_frete_senat +   #   8
            campo_valor_inss       +   #   8
            campo_pensao_alim      +   #  11
            campo_outros_desc      +   #  11
            campo_outros_prov      +   #  11
            campo_data_pagto       +   #   8
            campo_base_irrf        +   #  11
            campo_qtd_dep_ir       +   #   3
            campo_valor_ir             #   8
        )                              # = 266

    except Exception as e:
        log.append(f"ERRO ao montar registro do contrib {cod_contrib}: {e}")
        return None

    if len(registro) != 266:
        log.append(
            f"ERRO: Registro com tamanho {len(registro)} (esperado 266). "
            f"Cód empresa={codigo_empresa}, contrib={cod_contrib}"
        )
        return None

    return registro


# ==============================
# GERAÇÃO DO TXT (versão Streamlit)
# ==============================
def gerar_txt_streamlit(arquivo_bytes, log):
    try:
        meta = ler_planilha_rpa(io.BytesIO(arquivo_bytes), log)

        if meta is None:
            log.append("ERRO: Nenhum metadado/registro válido. Abortando.")
            return None, None

        meta["registros"].sort(
            key=lambda r: (
                str(r.get("cod_contrib", "")),
                excel_date_to_datetime(r.get("data_pagto"))
                    or ultimo_dia_competencia(meta["competencia"]),
                int(r.get("rpa_num")     or 0),
                int(r.get("linha_excel") or 0),
            )
        )

        linhas_txt = []
        acum_mes   = {}

        for reg in meta["registros"]:
            linha = montar_registro_lancamento(meta, reg, log, acum_mes)
            if linha is not None:
                linhas_txt.append(linha)

        if any(str(l).startswith("ERRO") for l in log):
            log.append("ERRO: Geração cancelada. TXT NÃO foi gerado.")
            return None, None

        if not linhas_txt:
            log.append("ERRO: Nenhum registro válido foi gerado para o TXT.")
            return None, None

        log.append(f"Arquivo TXT gerado com {len(linhas_txt)} registros.")
        return linhas_txt, meta

    except Exception:
        log.append("ERRO FATAL durante a geração do arquivo.")
        log.append(traceback.format_exc())
        return None, None


# ==============================
# INTERFACE STREAMLIT
# ==============================
def main():
    st.set_page_config(
        page_title=f"Domínio Sistemas | Thomson Reuters",
        page_icon="🟠",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    apply_tr_theme()

    st.markdown(
        f"""
        <div style="background:#444444; padding:24px 28px 18px 28px; border-radius:8px;
                    border-top:6px solid #FF8000; margin-bottom:28px;">
            <h2 style="color:#FF8000; margin:0; font-family:'Segoe UI',Arial,sans-serif;">
                🧾 Gerador de Arquivo TXT — RPA &nbsp;|&nbsp; {VERSAO}
            </h2>
            <p style="color:#DDDDDD; margin:6px 0 0 0; font-family:'Segoe UI',Arial,sans-serif;">
                Selecione o Excel de origem e clique em <strong>Gerar arquivo TXT</strong>.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if "log"          not in st.session_state:
        st.session_state.log          = [f"Aplicação pronta. Versão: {VERSAO}"]
    if "txt_gerado"   not in st.session_state:
        st.session_state.txt_gerado   = None
    if "nome_arquivo" not in st.session_state:
        st.session_state.nome_arquivo = "saida.txt"

    arquivo = st.file_uploader(
        "Excel de origem",
        type=["xlsx", "xls"],
        help="Selecione o arquivo Excel com os dados de RPA",
    )

    col1, col2 = st.columns([1, 1])

    with col1:
        gerar = st.button(
            "▶ Gerar arquivo TXT",
            disabled=(arquivo is None),
            use_container_width=True,
            type="primary",
        )
    with col2:
        limpar = st.button("🗑 Limpar", use_container_width=True)

    if limpar:
        st.session_state.log          = ["Campos limpos."]
        st.session_state.txt_gerado   = None
        st.session_state.nome_arquivo = "saida.txt"
        st.rerun()

    if gerar and arquivo is not None:
        st.session_state.log          = ["Iniciando geração do arquivo TXT..."]
        st.session_state.txt_gerado   = None
        st.session_state.nome_arquivo = "saida.txt"

        linhas, meta = gerar_txt_streamlit(arquivo.read(), st.session_state.log)

        if linhas and meta:
            conteudo = "\n".join(linhas) + "\n"
            st.session_state.txt_gerado   = conteudo.encode("latin-1", errors="replace")
            cod_emp     = str(meta["codigo_empresa"])
            competencia = competencia_aaaamm(meta["competencia"])
            st.session_state.nome_arquivo = f"{cod_emp}_RPA_competencia_{competencia}.txt"

        st.rerun()

    if st.session_state.txt_gerado is not None:
        st.success("✅ Arquivo gerado com sucesso!")
        st.download_button(
            label="⬇ Baixar arquivo TXT",
            data=st.session_state.txt_gerado,
            file_name=st.session_state.nome_arquivo,
            mime="text/plain",
            use_container_width=True,
            type="primary",
        )

    st.markdown("**Log de processamento**")
    log_texto = "\n".join(st.session_state.log)
    tem_erro  = any(str(l).startswith("ERRO") for l in st.session_state.log)
    cor_borda = "#D32F2F" if tem_erro else "#388E3C"

    st.markdown(
        f"""
        <div style="background:#FCFCFC; border:1px solid {cor_borda};
                    border-radius:6px; padding:14px;
                    font-family:Consolas,monospace; font-size:13px;
                    white-space:pre-wrap; max-height:340px;
                    overflow-y:auto; color:#1F1F1F;">
{log_texto}
        </div>
        """,
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
