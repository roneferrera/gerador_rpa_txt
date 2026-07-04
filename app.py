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
    if v < 0:
        return 0.0
    return v


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
    if len(s) > tamanho:
        s = s[-tamanho:]
    else:
        s = s.zfill(tamanho)
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
    if len(s) > tamanho:
        s = s[-tamanho:]
    else:
        s = s.zfill(tamanho)
    return s


def fmt_str(texto, tamanho):
    if texto is None:
        texto = ""
    try:
        if pd.isna(texto):
            texto = ""
    except Exception:
        pass
    texto = str(texto)
    return texto.ljust(tamanho)[:tamanho]


def competencia_aaaamm(data_excel):
    dt = excel_date_to_datetime(data_excel)
    if dt is None:
        return "000000"
    return dt.strftime("%Y%m")


def ultimo_dia_competencia(data_excel):
    dt = excel_date_to_datetime(data_excel)
    if dt is None:
        return None
    ano = dt.year
    mes = dt.month
    if mes == 12:
        prox = datetime(ano + 1, 1, 1)
    else:
        prox = datetime(ano, mes + 1, 1)
    ultimo = prox - pd.Timedelta(days=1)
    return ultimo


def tabela_ir_por_data_pagto(data_pagto_dt):
    if data_pagto_dt is None:
        return TABELA_IR_TRADICIONAL
    d = data_pagto_dt.date()
    return TABELA_IR_ATE_042025 if d < DATA_CORTE_TABELA_IR else TABELA_IR_TRADICIONAL


def deducao_simplificada_por_data_pagto(data_pagto_dt):
    if data_pagto_dt is None:
        return 0.0
    d = data_pagto_dt.date()
    return 564.80 if d < DATA_CORTE_TABELA_IR else 607.20


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
    if data_pagto_dt.year == 2025:
        return TETO_INSS_2025
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
        competencia
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
    aliquota = 0.0
    deducao = 0.0
    for limite, aliq, ded in tabela:
        if limite is None or base <= limite:
            aliquota = aliq
            deducao = ded
            break
    imposto_bruto = truncar(base * aliquota, casas=2)
    irrf = truncar(imposto_bruto - deducao, casas=2)
    if irrf < 0:
        irrf = 0.0
    return irrf


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
        parcela = truncar(0.133145 * rt, casas=2)
        return truncar(978.62 - parcela, casas=2)
    return 0.0


def calcular_irrf_2026_por_base(BC, rendimento_tributavel):
    if BC is None or BC <= 0:
        return 0.0
    ir_tabela = calcular_irrf_tabela(BC, TABELA_IR_TRADICIONAL)
    if ir_tabela <= 0:
        return 0.0
    red = reducao_mensal_2026(rendimento_tributavel)
    red_aplicada = min(red, ir_tabela)
    ir_final = truncar(ir_tabela - red_aplicada, casas=2)
    if ir_final < 0:
        ir_final = 0.0
    return ir_final


def calcular_irrf_mais_vantajoso_base100(base_bruta, dependentes, tabela, ded_simpl):
    if base_bruta is None or base_bruta <= 0:
        return 0.0, 0.0
    dep_int = 0 if (dependentes is None or pd.isna(dependentes)) else int(dependentes)
    red_dep = truncar(dep_int * VALOR_DEP, casas=2)
    base_geral = truncar(base_bruta - red_dep, casas=2)
    ir_geral = calcular_irrf_tabela(base_geral, tabela)
    base_simpl = truncar(base_bruta - ded_simpl, casas=2)
    ir_simpl = calcular_irrf_tabela(base_simpl, tabela)
    if (ir_simpl < ir_geral) or (ir_simpl == ir_geral and base_simpl <= base_geral):
        return ir_simpl, base_simpl
    return ir_geral, base_geral


def calcular_irrf_base60_legal(bruto, inss, dependentes, tabela):
    if bruto is None or bruto <= 0:
        return 0.0, 0.0
    base60 = truncar(bruto * 0.60, casas=2)
    dep_int = 0 if (dependentes is None or pd.isna(dependentes)) else int(dependentes)
    red_dep = truncar(dep_int * VALOR_DEP, casas=2)
    base = truncar(base60 - inss - red_dep, casas=2)
    ir = calcular_irrf_tabela(base, tabela)
    return ir, base


def calcular_irrf_base60_mais_vantajoso_2025(bruto, inss, dependentes, tabela, ded_simpl):
    if bruto is None or bruto <= 0:
        return 0.0, 0.0
    base60 = truncar(bruto * 0.60, casas=2)
    ir_geral, base_geral = calcular_irrf_base60_legal(bruto, inss, dependentes, tabela)
    base_simpl = truncar(base60 - ded_simpl, casas=2)
    ir_simpl = calcular_irrf_tabela(base_simpl, tabela)
    if (ir_simpl < ir_geral) or (ir_simpl == ir_geral and base_simpl <= base_geral):
        return ir_simpl, base_simpl
    return ir_geral, base_geral


def calcular_irrf_base10_legal(bruto, inss, dependentes, tabela):
    if bruto is None or bruto <= 0:
        return 0.0, 0.0
    base10 = truncar(bruto * 0.10, casas=2)
    dep_int = 0 if (dependentes is None or pd.isna(dependentes)) else int(dependentes)
    red_dep = truncar(dep_int * VALOR_DEP, casas=2)
    base = truncar(base10 - inss - red_dep, casas=2)
    ir = calcular_irrf_tabela(base, tabela)
    return ir, base


def calcular_irrf_base10_mais_vantajoso_2025(bruto, inss, dependentes, tabela, ded_simpl):
    if bruto is None or bruto <= 0:
        return 0.0, 0.0
    base10 = truncar(bruto * 0.10, casas=2)
    ir_geral, base_geral = calcular_irrf_base10_legal(bruto, inss, dependentes, tabela)
    base_simpl = truncar(base10 - ded_simpl, casas=2)
    ir_simpl = calcular_irrf_tabela(base_simpl, tabela)
    if (ir_simpl < ir_geral) or (ir_simpl == ir_geral and base_simpl <= base_geral):
        return ir_simpl, base_simpl
    return ir_geral, base_geral


def calcular_irrf_base60_mais_vantajoso_2026(bruto, inss, dependentes, ded_simpl, rendimento_tributavel):
    if bruto is None or bruto <= 0:
        return 0.0, 0.0
    base60 = truncar(bruto * 0.60, casas=2)
    dep_int = 0 if (dependentes is None or pd.isna(dependentes)) else int(dependentes)
    red_dep = truncar(dep_int * VALOR_DEP, casas=2)
    base_legal = truncar(base60 - inss - red_dep, casas=2)
    ir_legal = calcular_irrf_2026_por_base(base_legal, rendimento_tributavel)
    base_simpl = truncar(base60 - ded_simpl, casas=2)
    ir_simpl = calcular_irrf_2026_por_base(base_simpl, rendimento_tributavel)
    if (ir_simpl < ir_legal) or (ir_simpl == ir_legal and base_simpl <= base_legal):
        return ir_simpl, base_simpl
    return ir_legal, base_legal


def calcular_irrf_base10_mais_vantajoso_2026(bruto, inss, dependentes, ded_simpl, rendimento_tributavel):
    if bruto is None or bruto <= 0:
        return 0.0, 0.0
    base10 = truncar(bruto * 0.10, casas=2)
    dep_int = 0 if (dependentes is None or pd.isna(dependentes)) else int(dependentes)
    red_dep = truncar(dep_int * VALOR_DEP, casas=2)
    base_legal = truncar(base10 - inss - red_dep, casas=2)
    ir_legal = calcular_irrf_2026_por_base(base_legal, rendimento_tributavel)
    base_simpl = truncar(base10 - ded_simpl, casas=2)
    ir_simpl = calcular_irrf_2026_por_base(base_simpl, rendimento_tributavel)
    if (ir_simpl < ir_legal) or (ir_simpl == ir_legal and base_simpl <= base_legal):
        return ir_simpl, base_simpl
    return ir_legal, base_legal


def calcular_irrf_mais_vantajoso_2026_base100(base_bruta, dependentes, rendimento_tributavel, ded_simpl):
    if base_bruta is None or base_bruta <= 0:
        return 0.0, 0.0, "nenhum"
    dep_int = 0 if (dependentes is None or pd.isna(dependentes)) else int(dependentes)
    red_dep = truncar(dep_int * VALOR_DEP, casas=2)
    base_legal = truncar(base_bruta - red_dep, casas=2)
    ir_legal = calcular_irrf_2026_por_base(base_legal, rendimento_tributavel)
    base_simpl = truncar(base_bruta - ded_simpl, casas=2)
    ir_simpl = calcular_irrf_2026_por_base(base_simpl, rendimento_tributavel)
    if (ir_simpl < ir_legal) or (ir_simpl == ir_legal and base_simpl <= base_legal):
        return ir_simpl, base_simpl, "simplificada"
    return ir_legal, base_legal, "legal"


def calcular_irrf_acumulado_generico(
    rendimento_tributavel_acum,
    inss_dedutivel_acum,
    dependentes,
    ano_ir,
    tabela_ir,
    ded_simpl
):
    if rendimento_tributavel_acum is None or rendimento_tributavel_acum <= 0:
        return 0.0, 0.0
    dep_int = 0 if (dependentes is None or pd.isna(dependentes)) else int(dependentes)
    if dep_int < 0:
        dep_int = 0
    red_dep = truncar(dep_int * VALOR_DEP, casas=2)
    base_legal = truncar(rendimento_tributavel_acum - inss_dedutivel_acum - red_dep, casas=2)
    if base_legal < 0:
        base_legal = 0.0
    base_simpl = truncar(rendimento_tributavel_acum - ded_simpl, casas=2)
    if base_simpl < 0:
        base_simpl = 0.0
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

    codigo_empresa = None
    razao_social = None
    cnpj = None
    competencia = None

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

    if codigo_empresa is None or (isinstance(codigo_empresa, float) and pd.isna(codigo_empresa)):
        log.append("ERRO: 'Codigo Empresa' não encontrado.")
        return None
    if razao_social is None or (isinstance(razao_social, float) and pd.isna(razao_social)):
        log.append("ERRO: 'Razão Social' não encontrado.")
        return None
    if cnpj is None or (isinstance(cnpj, float) and pd.isna(cnpj)):
        log.append("ERRO: 'CNPJ' não encontrado.")
        return None
    if competencia is None or (isinstance(competencia, float) and pd.isna(competencia)):
        log.append("ERRO: 'Competencia' não encontrada.")
        return None

    codigo_empresa = int(codigo_empresa)

    inicio = None
    tem_cpf = False
    ncol = df.shape[1]

    for i in range(len(df)):
        def cell(r, c):
            if c >= ncol:
                return None
            return df.iloc[r, c]

        c0 = cell(i, 0)
        c1 = cell(i, 1)
        c2 = cell(i, 2)
        c3 = cell(i, 3)
        c4 = cell(i, 4)
        c5 = cell(i, 5)
        c6 = cell(i, 6)
        c7 = cell(i, 7)
        c13 = cell(i, 13)

        c0s = "" if (c0 is None or pd.isna(c0)) else str(c0).replace("RELAÇÃO DE RENDIMENTOS - RPA:", "").strip()
        c1s = "" if (c1 is None or pd.isna(c1)) else str(c1).strip()
        c2s = "" if (c2 is None or pd.isna(c2)) else str(c2).strip()
        c3s = "" if (c3 is None or pd.isna(c3)) else str(c3).strip()
        c4s = "" if (c4 is None or pd.isna(c4)) else str(c4).strip()
        c5s = "" if (c5 is None or pd.isna(c5)) else str(c5).strip()
        c6s = "" if (c6 is None or pd.isna(c6)) else str(c6).strip()
        c7s = "" if (c7 is None or pd.isna(c7)) else str(c7).strip()
        c13s = "" if (c13 is None or pd.isna(c13)) else str(c13).strip()

        if (
            c0s == "Código" and c1s == "Nome" and c2s == "CPF" and
            c3s == "Quantidade" and c4s == "Categoria" and c5s == "Próxima" and
            c6s == "Descrição" and c7s == "Rendimento" and c13s == "Data ISS"
        ):
            inicio = i + 2
            tem_cpf = True
            break

        if (
            c0s == "Código" and c1s == "Nome" and
            c2s == "Quantidade" and c3s == "Categoria" and
            c4s == "Próxima" and c5s == "Descrição" and c6s == "Rendimento"
        ):
            inicio = i + 2
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
        s = str(v).strip()
        if not s:
            return 0.0
        s = re.sub(r"[^0-9,\.\-]", "", s)
        if s in ("", "-", ",", ".", "-.", "-,"):
            return 0.0
        if "." in s and "," in s:
            s = s.replace(".", "").replace(",", ".")
        else:
            if "," in s:
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
        linha = df.iloc[i]
        cod_contrib = linha[0] if len(linha) > 0 else None
        if cod_contrib is None or pd.isna(cod_contrib):
            continue
        try:
            if tem_cpf:
                nome = linha[1]
                dependentes = linha[3]
                esocial = linha[4]
                rpa_num = linha[5]
                atividade = linha[6]
                bruto = linha[7]
                data_pagto = linha[8]
                pensao = linha[9]
                outros_desc = linha[10]
                outros_prov = linha[11]
                perc_iss = linha[12]
                data_iss = linha[13]
            else:
                nome = linha[1]
                dependentes = linha[2]
                esocial = linha[3]
                rpa_num = linha[4]
                atividade = linha[5]
                bruto = linha[6]
                data_pagto = linha[7]
                pensao = linha[8]
                outros_desc = linha[9]
                outros_prov = linha[10]
                perc_iss = linha[11]
                data_iss = linha[12]

            if bruto is None or pd.isna(bruto):
                log.append(f"Aviso: linha {i+1} sem BRUTO. Código: {cod_contrib}. Pulando.")
                continue

            registros.append({
                "cod_contrib": cod_contrib,
                "nome": nome,
                "dependentes": dependentes,
                "esocial": esocial,
                "rpa_num": rpa_num,
                "atividade": atividade,
                "bruto": _num_or_zero(bruto),
                "data_pagto": data_pagto,
                "pensao_alim": _num_or_zero(pensao),
                "outros_desc": _num_or_zero(outros_desc),
                "outros_prov": _num_or_zero(outros_prov),
                "perc_iss": _num_or_zero(perc_iss),
                "valor_iss": 0.0,
                "data_iss": data_iss,
                "linha_excel": i + 1,
            })
        except Exception as e:
            log.append(f"ERRO ao ler linha {i+1}: {e}")

    return {
        "codigo_empresa": codigo_empresa,
        "razao_social": razao_social,
        "cnpj": str(cnpj),
        "competencia": competencia,
        "registros": registros,
    }


# ==============================
# MONTAGEM DO REGISTRO TXT (266)
# ==============================
def montar_registro_lancamento(meta, reg, log, acum_mes):
    codigo_empresa = meta["codigo_empresa"]
    competencia_data = meta["competencia"]
    competencia_str = competencia_aaaamm(competencia_data)

    data_pagto_excel = reg.get("data_pagto")
    data_pagto_dt = excel_date_to_datetime(data_pagto_excel)

    if data_pagto_dt is None:
        data_pagto_dt = ultimo_dia_competencia(competencia_data)

    data_pagto_str = "00000000" if data_pagto_dt is None else data_pagto_dt.strftime("%Y%m%d")
    ano_ir = data_pagto_dt.year if data_pagto_dt is not None else None

    tabela_ir = tabela_ir_por_data_pagto(data_pagto_dt)
    ded_simpl = deducao_simplificada_por_data_pagto_ou_ano(data_pagto_dt)

    cod_contrib = reg["cod_contrib"]
    dependentes = reg["dependentes"]
    rpa_num = reg["rpa_num"]
    atividade = reg["atividade"]
    bruto = limpar_negativo(reg["bruto"])

    perc_iss = limpar_negativo(reg.get("perc_iss", 0.0))
    pensao_alim = limpar_negativo(reg.get("pensao_alim", 0.0))
    outros_desc = limpar_negativo(reg.get("outros_desc", 0.0))
    outros_prov = limpar_negativo(reg.get("outros_prov", 0.0))

    data_iss_excel = reg.get("data_iss")
    dt_iss = excel_date_to_datetime(data_iss_excel)
    data_venc_iss = "00000000" if dt_iss is None else dt_iss.strftime("%Y%m%d")

    esocial = reg.get("esocial")
    try:
        esocial_int = int(esocial) if not pd.isna(esocial) else None
    except Exception:
        esocial_int = None

    chave = chave_acumulacao_mes(meta, reg, data_pagto_dt)

    if chave not in acum_mes:
        acum_mes[chave] = {
            "base_inss_empresa": 0.0,
            "inss_retido_empresa": 0.0,
            "outras_fontes_base": 0.0,
            "rend_trib_irrf": 0.0,
            "inss_dedutivel_irrf": 0.0,
            "irrf_retido": 0.0,
            "dependentes": 0,
        }

    ac = acum_mes[chave]

    inss_frete_sest = 0.0
    inss_frete_senat = 0.0
    base_irrf = 0.0
    ir_calculado = 0.0

    base_inss_registro_original = bruto
    aliquota_inss = 0.11

    if esocial_int in (712, 734):
        base_inss_registro_original = truncar(bruto * 0.20, casas=2)
        aliquota_inss = 0.20 if esocial_int == 734 else 0.11
        inss_frete_sest = truncar(base_inss_registro_original * 0.015, casas=2)
        inss_frete_senat = truncar(base_inss_registro_original * 0.010, casas=2)

    teto_inss = teto_inss_por_data_pagto(data_pagto_dt)

    outras_fontes_base = truncar(ac.get("outras_fontes_base", 0.0), casas=2)
    if outras_fontes_base < 0:
        outras_fontes_base = 0.0

    saldo_teto = truncar(teto_inss - outras_fontes_base, casas=2)
    if saldo_teto < 0:
        saldo_teto = 0.0

    base_empresa_anterior = truncar(ac["base_inss_empresa"], casas=2)
    base_empresa_nova = truncar(base_empresa_anterior + base_inss_registro_original, casas=2)

    base_limitada_anterior = min(base_empresa_anterior, saldo_teto)
    base_limitada_nova = min(base_empresa_nova, saldo_teto)

    base_inss_registro_limitada = truncar(base_limitada_nova - base_limitada_anterior, casas=2)
    if base_inss_registro_limitada < 0:
        base_inss_registro_limitada = 0.0

    inss = truncar(base_inss_registro_limitada * aliquota_inss, casas=2)
    if inss < 0:
        inss = 0.0

    ac["base_inss_empresa"] = base_empresa_nova
    ac["inss_retido_empresa"] = truncar(ac["inss_retido_empresa"] + inss, casas=2)

    base_inss = base_inss_registro_limitada

    rendimento_tributavel_registro = obter_rendimento_tributavel_irrf(bruto, esocial_int)

    dep_out = 0 if (dependentes is None or pd.isna(dependentes)) else int(dependentes)
    if dep_out < 0:
        dep_out = 0

    deduz_inss = esocial_int in (711, 712)

    ac["rend_trib_irrf"] = truncar(ac["rend_trib_irrf"] + rendimento_tributavel_registro, casas=2)
    ac["inss_dedutivel_irrf"] = truncar(ac["inss_dedutivel_irrf"] + inss, casas=2)
    ac["dependentes"] = max(ac["dependentes"], dep_out)

    rendimento_tributavel_acum = ac["rend_trib_irrf"]
    inss_dedutivel_acum = ac["inss_dedutivel_irrf"] if deduz_inss else 0.0
    dependentes_acum = ac["dependentes"]

    if ano_ir == 2025:
        ir_total_mes, base_irrf_mes = calcular_irrf_acumulado_generico(
            rendimento_tributavel_acum=rendimento_tributavel_acum,
            inss_dedutivel_acum=inss_dedutivel_acum,
            dependentes=dependentes_acum,
            ano_ir=2025,
            tabela_ir=tabela_ir,
            ded_simpl=ded_simpl
        )
    elif ano_ir == 2026:
        ir_total_mes, base_irrf_mes = calcular_irrf_acumulado_generico(
            rendimento_tributavel_acum=rendimento_tributavel_acum,
            inss_dedutivel_acum=inss_dedutivel_acum,
            dependentes=dependentes_acum,
            ano_ir=2026,
            tabela_ir=tabela_ir,
            ded_simpl=ded_simpl
        )
    else:
        log.append(
            f"Aviso: ano de pagamento desconhecido ({ano_ir}) para contrib "
            f"{cod_contrib}; usando regra 2025."
        )
        ir_total_mes, base_irrf_mes = calcular_irrf_acumulado_generico(
            rendimento_tributavel_acum=rendimento_tributavel_acum,
            inss_dedutivel_acum=inss_dedutivel_acum,
            dependentes=dependentes_acum,
            ano_ir=2025,
            tabela_ir=tabela_ir,
            ded_simpl=ded_simpl
        )

    irrf_ja_retido = truncar(ac["irrf_retido"], casas=2)
    ir_calculado = truncar(ir_total_mes - irrf_ja_retido, casas=2)
    if ir_calculado < 0:
        ir_calculado = 0.0

    ac["irrf_retido"] = truncar(ac["irrf_retido"] + ir_calculado, casas=2)

    base_irrf = base_irrf_mes

    if perc_iss and float(perc_iss) != 0.0:
        valor_iss = truncar(bruto * (perc_iss / 100.0), casas=2)
    else:
        perc_iss = 0.0
        valor_iss = 0.0

    valor_iss = limpar_negativo(valor_iss)
    base_inss = limpar_negativo(base_inss)
    inss_frete_sest = limpar_negativo(inss_frete_sest)
    inss_frete_senat = limpar_negativo(inss_frete_senat)
    inss = limpar_negativo(inss)
    base_irrf = limpar_negativo(base_irrf)
    ir_calculado = limpar_negativo(ir_calculado)

    try:
        campo_codigo_empresa   = fmt_int(codigo_empresa, 7)
        campo_codigo_contrib   = fmt_int(cod_contrib, 10)
        campo_competencia      = competencia_str
        campo_desc_atividade   = fmt_str(atividade, 100)
        campo_num_rpa          = fmt_int(rpa_num, 10)
        campo_rendimento_bruto = fmt_num(bruto, 11, casas=2, permitir_negativo=False)
        campo_percentual_iss   = fmt_num(perc_iss, 5, casas=2, permitir_negativo=False)
        campo_valor_iss        = fmt_num(valor_iss, 11, casas=2, permitir_negativo=False)
        campo_data_venc_iss    = data_venc_iss
        campo_base_inss        = fmt_num(base_inss, 11, casas=2, permitir_negativo=False)
        campo_inss_frete_sest  = fmt_num(inss_frete_sest, 8, casas=2, permitir_negativo=False)
        campo_inss_frete_senat = fmt_num(inss_frete_senat, 8, casas=2, permitir_negativo=False)
        campo_valor_inss       = fmt_num(inss, 8, casas=2, permitir_negativo=False)
        campo_pensao_alim      = fmt_num(pensao_alim, 11, casas=2, permitir_negativo=False)
        campo_outros_desc      = fmt_num(outros_desc, 11, casas=2, permitir_negativo=False)
        campo_outros_prov      = fmt_num(outros_prov, 11, casas=2, permitir_negativo=False)
        campo_data_pagto       = data_pagto_str
        campo_base_irrf        = fmt_num(base_irrf, 11, casas=2, permitir_negativo=False)
        campo_qtd_dep_ir       = fmt_int(dep_out, 3)
        campo_valor_ir         = fmt_num(ir_calculado, 8, casas=2, permitir_negativo=False)

        registro = (
            campo_codigo_empresa +
            campo_codigo_contrib +
            campo_competencia +
            campo_desc_atividade +
            campo_num_rpa +
            campo_rendimento_bruto +
            campo_percentual_iss +
            campo_valor_iss +
            campo_data_venc_iss +
            campo_base_inss +
            campo_inss_frete_sest +
            campo_inss_frete_senat +
            campo_valor_inss +
            campo_pensao_alim +
            campo_outros_desc +
            campo_outros_prov +
            campo_data_pagto +
            campo_base_irrf +
            campo_qtd_dep_ir +
            campo_valor_ir
        )

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
        excel_buffer = io.BytesIO(arquivo_bytes)
        meta = ler_planilha_rpa(excel_buffer, log)

        if meta is None:
            log.append("ERRO: Nenhum metadado/registro válido. Abortando.")
            return None

        meta["registros"].sort(
            key=lambda r: (
                str(r.get("cod_contrib", "")),
                excel_date_to_datetime(r.get("data_pagto"))
                    or ultimo_dia_competencia(meta["competencia"]),
                int(r.get("rpa_num") or 0),
                int(r.get("linha_excel") or 0),
            )
        )

        linhas_txt = []
        acum_mes = {}

        for reg in meta["registros"]:
            linha = montar_registro_lancamento(meta, reg, log, acum_mes)
            if linha is not None:
                linhas_txt.append(linha)

        if any(str(l).startswith("ERRO") for l in log):
            log.append("ERRO: Geração cancelada. TXT NÃO foi gerado.")
            return None

        if len(linhas_txt) == 0:
            log.append("ERRO: Nenhum registro válido foi gerado para o TXT.")
            return None

        log.append(f"Arquivo TXT gerado com {len(linhas_txt)} registros.")
        return linhas_txt

    except Exception:
        log.append("ERRO FATAL durante a geração do arquivo.")
        log.append(traceback.format_exc())
        return None


# ==============================
# CSS — Tema Thomson Reuters / Domínio Sistemas (escuro)
# ==============================
TR_DARK_CSS = """
<style>
/* ── Paleta ──────────────────────────────────────────────
   Fundo principal  : #0D1117
   Superfície card  : #161B22
   Borda sutil      : #21262D
   Accent TR laranja: #FF6200
   Accent TR azul   : #006AFF
   Texto primário   : #E6EDF3
   Texto secundário : #8B949E
   Sucesso          : #3FB950
   Erro             : #F85149
   Aviso            : #D29922
──────────────────────────────────────────────────────── */

/* Fundo global */
[data-testid="stAppViewContainer"],
[data-testid="stMain"],
.main {
    background-color: #0D1117 !important;
}

[data-testid="stSidebar"] {
    background-color: #0D1117 !important;
    border-right: 1px solid #21262D !important;
}

/* Texto base */
html, body, [class*="css"], p, span, label, div {
    color: #E6EDF3 !important;
    font-family: 'Segoe UI', 'Inter', 'Helvetica Neue', Arial, sans-serif !important;
}

/* Títulos */
h1, h2, h3, h4, h5, h6 {
    color: #E6EDF3 !important;
    font-weight: 600 !important;
    letter-spacing: -0.02em !important;
}

/* ── Botão primário (Gerar) ─────────────────────────── */
[data-testid="stButton"] button[kind="primary"],
button[kind="primary"] {
    background: linear-gradient(135deg, #FF6200 0%, #E05500 100%) !important;
    color: #FFFFFF !important;
    border: none !important;
    border-radius: 6px !important;
    font-weight: 600 !important;
    font-size: 14px !important;
    letter-spacing: 0.02em !important;
    padding: 10px 20px !important;
    transition: opacity 0.2s ease, transform 0.1s ease !important;
    box-shadow: 0 2px 8px rgba(255,98,0,0.35) !important;
}
[data-testid="stButton"] button[kind="primary"]:hover {
    opacity: 0.88 !important;
    transform: translateY(-1px) !important;
}
[data-testid="stButton"] button[kind="primary"]:active {
    transform: translateY(0px) !important;
}
[data-testid="stButton"] button[kind="primary"]:disabled {
    background: #2D333B !important;
    color: #484F58 !important;
    box-shadow: none !important;
    cursor: not-allowed !important;
}

/* ── Botão secundário (Limpar / Download) ──────────── */
[data-testid="stButton"] button[kind="secondary"],
button[kind="secondary"] {
    background-color: #21262D !important;
    color: #E6EDF3 !important;
    border: 1px solid #30363D !important;
    border-radius: 6px !important;
    font-weight: 500 !important;
    font-size: 14px !important;
    transition: background-color 0.2s ease, border-color 0.2s ease !important;
}
[data-testid="stButton"] button[kind="secondary"]:hover {
    background-color: #2D333B !important;
    border-color: #484F58 !important;
}

/* ── Download button ────────────────────────────────── */
[data-testid="stDownloadButton"] button {
    background: linear-gradient(135deg, #006AFF 0%, #0054CC 100%) !important;
    color: #FFFFFF !important;
    border: none !important;
    border-radius: 6px !important;
    font-weight: 600 !important;
    font-size: 14px !important;
    box-shadow: 0 2px 8px rgba(0,106,255,0.35) !important;
    transition: opacity 0.2s ease !important;
}
[data-testid="stDownloadButton"] button:hover {
    opacity: 0.88 !important;
}

/* ── File uploader ──────────────────────────────────── */
[data-testid="stFileUploader"] {
    background-color: #161B22 !important;
    border: 1.5px dashed #30363D !important;
    border-radius: 8px !important;
    padding: 8px !important;
    transition: border-color 0.2s ease !important;
}
[data-testid="stFileUploader"]:hover {
    border-color: #FF6200 !important;
}
[data-testid="stFileUploaderDropzoneInstructions"] div,
[data-testid="stFileUploaderDropzone"] label,
[data-testid="stFileUploaderDropzone"] span {
    color: #8B949E !important;
}

/* ── Alertas / st.success ───────────────────────────── */
[data-testid="stAlert"] {
    border-radius: 8px !important;
    border-left-width: 4px !important;
}
[data-testid="stAlert"][data-baseweb="notification"][kind="positive"],
div[data-testid="stAlert"].st-emotion-cache-1xarl3l {
    background-color: #0D2818 !important;
    border-left-color: #3FB950 !important;
    color: #3FB950 !important;
}

/* ── Scrollbar ──────────────────────────────────────── */
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: #161B22; }
::-webkit-scrollbar-thumb { background: #30363D; border-radius: 3px; }
::-webkit-scrollbar-thumb:hover { background: #484F58; }

/* ── Divisor ────────────────────────────────────────── */
hr {
    border-color: #21262D !important;
    margin: 20px 0 !important;
}
</style>
"""


# ==============================
# INTERFACE STREAMLIT — Tema TR/Domínio Escuro
# ==============================
def main():
    st.set_page_config(
        page_title=f"Gerador TXT — RPA | {VERSAO}",
        page_icon="📄",
        layout="centered",
    )

    # Injeta CSS do tema
    st.markdown(TR_DARK_CSS, unsafe_allow_html=True)

    # ── Header ──────────────────────────────────────────
    st.markdown(
        f"""
        <div style="
            background: linear-gradient(135deg, #161B22 0%, #1C2128 100%);
            border: 1px solid #21262D;
            border-top: 3px solid #FF6200;
            border-radius: 10px;
            padding: 28px 28px 20px 28px;
            margin-bottom: 28px;
        ">
            <div style="display:flex; align-items:center; gap:14px; margin-bottom:6px;">
                <span style="font-size:28px;">📄</span>
                <div>
                    <div style="
                        font-size:11px;
                        font-weight:700;
                        letter-spacing:0.12em;
                        color:#FF6200;
                        text-transform:uppercase;
                        margin-bottom:2px;
                    ">Thomson Reuters · Domínio Sistemas</div>
                    <h2 style="
                        margin:0;
                        font-size:22px;
                        font-weight:700;
                        color:#E6EDF3;
                        letter-spacing:-0.02em;
                    ">Gerador de Arquivo TXT — RPA</h2>
                </div>
                <div style="
                    margin-left:auto;
                    background:#21262D;
                    border:1px solid #30363D;
                    border-radius:20px;
                    padding:3px 12px;
                    font-size:12px;
                    font-weight:600;
                    color:#8B949E;
                    letter-spacing:0.06em;
                ">{VERSAO}</div>
            </div>
            <p style="
                margin:10px 0 0 42px;
                font-size:13.5px;
                color:#8B949E;
                line-height:1.5;
            ">Importe o Excel de origem e clique em <strong style="color:#E6EDF3;">Gerar arquivo TXT</strong> para processar os lançamentos RPA.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ── Session state ────────────────────────────────────
    if "log" not in st.session_state:
        st.session_state.log = [f"✔ Aplicação pronta — Versão {VERSAO}"]
    if "txt_gerado" not in st.session_state:
        st.session_state.txt_gerado = None
    if "meta_info" not in st.session_state:
        st.session_state.meta_info = None

    # ── Upload card ──────────────────────────────────────
    st.markdown(
        """
        <div style="
            background:#161B22;
            border:1px solid #21262D;
            border-radius:8px;
            padding:20px 20px 4px 20px;
            margin-bottom:16px;
        ">
            <div style="
                font-size:11px;
                font-weight:700;
                letter-spacing:0.1em;
                color:#8B949E;
                text-transform:uppercase;
                margin-bottom:10px;
            ">📂 Arquivo de entrada</div>
        """,
        unsafe_allow_html=True,
    )

    arquivo = st.file_uploader(
        "Selecione o arquivo Excel (.xlsx / .xls)",
        type=["xlsx", "xls"],
        help="Planilha de Relação de Rendimentos — RPA",
        label_visibility="collapsed",
    )

    st.markdown("</div>", unsafe_allow_html=True)

    # ── Botões ───────────────────────────────────────────
    col1, col2 = st.columns([3, 1])

    with col1:
        gerar = st.button(
            "▶  Gerar arquivo TXT",
            disabled=(arquivo is None),
            use_container_width=True,
            type="primary",
        )

    with col2:
        limpar = st.button(
            "🗑  Limpar",
            use_container_width=True,
            type="secondary",
        )

    if limpar:
        st.session_state.log = ["🗑 Campos limpos. Pronto para nova operação."]
        st.session_state.txt_gerado = None
        st.session_state.meta_info = None
        st.rerun()

    # ── Processamento ────────────────────────────────────
    if gerar and arquivo is not None:
        st.session_state.log = ["⏳ Iniciando processamento do arquivo..."]
        st.session_state.txt_gerado = None
        st.session_state.meta_info = None

        arquivo_bytes = arquivo.read()

        with st.spinner("Processando lançamentos RPA…"):
            linhas = gerar_txt_streamlit(arquivo_bytes, st.session_state.log)

        if linhas:
            conteudo = "\n".join(linhas) + "\n"
            st.session_state.txt_gerado = conteudo.encode("latin-1", errors="replace")

            # Captura metadados para o painel de resumo
            try:
                import io as _io
                _buf = _io.BytesIO(arquivo_bytes)
                _meta_tmp = ler_planilha_rpa(_buf, [])
                if _meta_tmp:
                    st.session_state.meta_info = {
                        "empresa": _meta_tmp.get("razao_social", "—"),
                        "cnpj": _meta_tmp.get("cnpj", "—"),
                        "competencia": competencia_aaaamm(_meta_tmp.get("competencia")),
                        "registros": len(linhas),
                    }
            except Exception:
                pass

        st.rerun()

    # ── Painel de resultado ──────────────────────────────
    if st.session_state.txt_gerado is not None:

        # Cards de resumo
        if st.session_state.meta_info:
            m = st.session_state.meta_info
            comp_fmt = (
                f"{m['competencia'][4:6]}/{m['competencia'][:4]}"
                if len(m["competencia"]) == 6
                else m["competencia"]
            )
            st.markdown(
                f"""
                <div style="
                    display:grid;
                    grid-template-columns: repeat(4, 1fr);
                    gap:10px;
                    margin: 20px 0 16px 0;
                ">
                    <div style="background:#161B22;border:1px solid #21262D;
                                border-top:2px solid #FF6200;border-radius:8px;padding:14px;">
                        <div style="font-size:10px;font-weight:700;letter-spacing:0.1em;
                                    color:#8B949E;text-transform:uppercase;margin-bottom:4px;">Empresa</div>
                        <div style="font-size:13px;font-weight:600;color:#E6EDF3;
                                    white-space:nowrap;overflow:hidden;text-overflow:ellipsis;"
                             title="{m['empresa']}">{m['empresa']}</div>
                    </div>
                    <div style="background:#161B22;border:1px solid #21262D;
                                border-top:2px solid #006AFF;border-radius:8px;padding:14px;">
                        <div style="font-size:10px;font-weight:700;letter-spacing:0.1em;
                                    color:#8B949E;text-transform:uppercase;margin-bottom:4px;">CNPJ</div>
                        <div style="font-size:13px;font-weight:600;color:#E6EDF3;">{m['cnpj']}</div>
                    </div>
                    <div style="background:#161B22;border:1px solid #21262D;
                                border-top:2px solid #D29922;border-radius:8px;padding:14px;">
                        <div style="font-size:10px;font-weight:700;letter-spacing:0.1em;
                                    color:#8B949E;text-transform:uppercase;margin-bottom:4px;">Competência</div>
                        <div style="font-size:13px;font-weight:600;color:#E6EDF3;">{comp_fmt}</div>
                    </div>
                    <div style="background:#161B22;border:1px solid #21262D;
                                border-top:2px solid #3FB950;border-radius:8px;padding:14px;">
                        <div style="font-size:10px;font-weight:700;letter-spacing:0.1em;
                                    color:#8B949E;text-transform:uppercase;margin-bottom:4px;">Registros</div>
                        <div style="font-size:20px;font-weight:700;color:#3FB950;">{m['registros']}</div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        # Alerta de sucesso customizado
        st.markdown(
            """
            <div style="
                background:#0D2818;
                border:1px solid #238636;
                border-left:4px solid #3FB950;
                border-radius:8px;
                padding:14px 18px;
                margin-bottom:14px;
                display:flex;
                align-items:center;
                gap:10px;
            ">
                <span style="font-size:18px;">✅</span>
                <span style="font-size:14px;font-weight:600;color:#3FB950;">
                    Arquivo gerado com sucesso! Clique no botão abaixo para baixar.
                </span>
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.download_button(
            label="⬇  Baixar arquivo TXT",
            data=st.session_state.txt_gerado,
            file_name="saida_rpa.txt",
            mime="text/plain",
            use_container_width=True,
            type="primary",
        )

    # ── Log de processamento ─────────────────────────────
    st.markdown("<div style='margin-top:28px;'></div>", unsafe_allow_html=True)

    tem_erro = any(str(l).startswith("ERRO") for l in st.session_state.log)
    tem_aviso = any(str(l).startswith("Aviso") for l in st.session_state.log)

    if tem_erro:
        cor_accent = "#F85149"
        cor_borda  = "#DA3633"
        cor_fundo  = "#1A0D0D"
        icone_log  = "🔴"
        label_log  = "Log de processamento — Erros detectados"
    elif tem_aviso:
        cor_accent = "#D29922"
        cor_borda  = "#9E6A03"
        cor_fundo  = "#1A1500"
        icone_log  = "🟡"
        label_log  = "Log de processamento — Avisos"
    else:
        cor_accent = "#3FB950"
        cor_borda  = "#238636"
        cor_fundo  = "#0D1A10"
        icone_log  = "🟢"
        label_log  = "Log de processamento"

    st.markdown(
        f"""
        <div style="
            background:#161B22;
            border:1px solid #21262D;
            border-top:2px solid {cor_accent};
            border-radius:8px;
            overflow:hidden;
        ">
            <div style="
                background:#1C2128;
                padding:10px 16px;
                border-bottom:1px solid #21262D;
                display:flex;
                align-items:center;
                gap:8px;
            ">
                <span>{icone_log}</span>
                <span style="font-size:12px;font-weight:700;letter-spacing:0.08em;
                             color:#8B949E;text-transform:uppercase;">{label_log}</span>
            </div>
            <div style="
                background:{cor_fundo};
                border-left:3px solid {cor_borda};
                margin:12px;
                border-radius:4px;
                padding:12px 14px;
                font-family:'Cascadia Code','Fira Code','Consolas',monospace;
                font-size:12.5px;
                line-height:1.7;
                white-space:pre-wrap;
                max-height:300px;
                overflow-y:auto;
                color:#C9D1D9;
            ">{chr(10).join(st.session_state.log)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ── Rodapé ───────────────────────────────────────────
    st.markdown(
        f"""
        <div style="
            margin-top:36px;
            padding-top:16px;
            border-top:1px solid #21262D;
            display:flex;
            justify-content:space-between;
            align-items:center;
        ">
            <span style="font-size:11px;color:#484F58;">
                Thomson Reuters · Domínio Sistemas · Gerador RPA
            </span>
            <span style="
                font-size:11px;
                font-weight:700;
                letter-spacing:0.08em;
                color:#FF6200;
            ">{VERSAO}</span>
        </div>
        """,
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
