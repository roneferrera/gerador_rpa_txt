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
# VERSÃO E CONSTANTES GLOBAIS
# ==============================
VERSAO           = "V1"
NOME_ARQUIVO_BGR = "Relação de Rendimentos - RPA.bgr"
NOME_ARQUIVO_TXT = "Relação de Rendimentos - RPA.txt"

# ==============================
# TABELAS FISCAIS
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
VALOR_DEP                 = 189.59
DATA_CORTE_TABELA_IR      = date(2025, 5, 1)
DEDUCAO_SIMPLIFICADA_2026 = 607.20
TETO_INSS_2025            = 8157.41
TETO_INSS_2026            = 8475.55


# ==============================
# FUNÇÕES AUXILIARES
# ==============================
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
    return math.floor(float(valor) * (10 ** casas)) / (10 ** casas)


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
    return s[-tamanho:] if len(s) > tamanho else s.zfill(tamanho)


def fmt_int(valor, tamanho):
    if valor is None:
        valor = 0
    try:
        if pd.isna(valor):
            valor = 0
    except Exception:
        pass
    s = f"{int(valor):d}"
    return s[-tamanho:] if len(s) > tamanho else s.zfill(tamanho)


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
    return "000000" if dt is None else dt.strftime("%Y%m")


def ultimo_dia_competencia(data_excel):
    dt = excel_date_to_datetime(data_excel)
    if dt is None:
        return None
    ano, mes = dt.year, dt.month
    prox = datetime(ano + 1, 1, 1) if mes == 12 else datetime(ano, mes + 1, 1)
    return prox - pd.Timedelta(days=1)


def tabela_ir_por_data_pagto(dt):
    if dt is None:
        return TABELA_IR_TRADICIONAL
    return TABELA_IR_ATE_042025 if dt.date() < DATA_CORTE_TABELA_IR else TABELA_IR_TRADICIONAL


def deducao_simplificada_por_data_pagto(dt):
    if dt is None:
        return 0.0
    return 564.80 if dt.date() < DATA_CORTE_TABELA_IR else 607.20


def deducao_simplificada_por_data_pagto_ou_ano(dt):
    if dt is None:
        return 0.0
    if dt.year >= 2026:
        return DEDUCAO_SIMPLIFICADA_2026
    return deducao_simplificada_por_data_pagto(dt)


def teto_inss_por_data_pagto(dt):
    if dt is None:
        return TETO_INSS_2026
    return TETO_INSS_2026 if dt.year >= 2026 else TETO_INSS_2025


def chave_acumulacao_mes(meta, reg, dt):
    comp = dt.strftime("%Y%m") if dt is not None else competencia_aaaamm(meta["competencia"])
    return (int(meta["codigo_empresa"]), str(reg["cod_contrib"]).strip(), comp)


def obter_rendimento_tributavel_irrf(bruto, esocial_int):
    bruto = limpar_negativo(bruto)
    if bruto <= 0:
        return 0.0
    if esocial_int in (711, 731, 734):
        return truncar(bruto * 0.60)
    if esocial_int == 712:
        return truncar(bruto * 0.10)
    return truncar(bruto)


def calcular_irrf_tabela(base, tabela):
    if base is None or base <= 0:
        return 0.0
    aliq, ded = 0.0, 0.0
    for limite, a, d in tabela:
        if limite is None or base <= limite:
            aliq, ded = a, d
            break
    return max(truncar(truncar(base * aliq) - ded), 0.0)


def reducao_mensal_2026(rt):
    if rt is None:
        return 0.0
    try:
        rt = float(rt)
    except Exception:
        return 0.0
    if rt <= 0:
        return 0.0
    if rt <= 5000.00:
        return 312.89
    if rt <= 7350.00:
        return truncar(978.62 - truncar(0.133145 * rt))
    return 0.0


def calcular_irrf_2026_por_base(BC, rt):
    if BC is None or BC <= 0:
        return 0.0
    ir = calcular_irrf_tabela(BC, TABELA_IR_TRADICIONAL)
    if ir <= 0:
        return 0.0
    red = reducao_mensal_2026(rt)
    return max(truncar(ir - min(red, ir)), 0.0)


def calcular_irrf_mais_vantajoso_base100(base_bruta, dependentes, tabela, ded_simpl):
    if base_bruta is None or base_bruta <= 0:
        return 0.0, 0.0
    dep_int    = 0 if (dependentes is None or pd.isna(dependentes)) else int(dependentes)
    red_dep    = truncar(dep_int * VALOR_DEP)
    base_geral = truncar(base_bruta - red_dep)
    ir_geral   = calcular_irrf_tabela(base_geral, tabela)
    base_simpl = truncar(base_bruta - ded_simpl)
    ir_simpl   = calcular_irrf_tabela(base_simpl, tabela)
    if (ir_simpl < ir_geral) or (ir_simpl == ir_geral and base_simpl <= base_geral):
        return ir_simpl, base_simpl
    return ir_geral, base_geral


def calcular_irrf_base60_legal(bruto, inss, dependentes, tabela):
    if bruto is None or bruto <= 0:
        return 0.0, 0.0
    base60  = truncar(bruto * 0.60)
    dep_int = 0 if (dependentes is None or pd.isna(dependentes)) else int(dependentes)
    red_dep = truncar(dep_int * VALOR_DEP)
    base    = truncar(base60 - inss - red_dep)
    return calcular_irrf_tabela(base, tabela), base


def calcular_irrf_base60_mais_vantajoso_2025(bruto, inss, dependentes, tabela, ded_simpl):
    if bruto is None or bruto <= 0:
        return 0.0, 0.0
    base60               = truncar(bruto * 0.60)
    ir_geral, base_geral = calcular_irrf_base60_legal(bruto, inss, dependentes, tabela)
    base_simpl           = truncar(base60 - ded_simpl)
    ir_simpl             = calcular_irrf_tabela(base_simpl, tabela)
    if (ir_simpl < ir_geral) or (ir_simpl == ir_geral and base_simpl <= base_geral):
        return ir_simpl, base_simpl
    return ir_geral, base_geral


def calcular_irrf_base10_legal(bruto, inss, dependentes, tabela):
    if bruto is None or bruto <= 0:
        return 0.0, 0.0
    base10  = truncar(bruto * 0.10)
    dep_int = 0 if (dependentes is None or pd.isna(dependentes)) else int(dependentes)
    red_dep = truncar(dep_int * VALOR_DEP)
    base    = truncar(base10 - inss - red_dep)
    return calcular_irrf_tabela(base, tabela), base


def calcular_irrf_base10_mais_vantajoso_2025(bruto, inss, dependentes, tabela, ded_simpl):
    if bruto is None or bruto <= 0:
        return 0.0, 0.0
    base10               = truncar(bruto * 0.10)
    ir_geral, base_geral = calcular_irrf_base10_legal(bruto, inss, dependentes, tabela)
    base_simpl           = truncar(base10 - ded_simpl)
    ir_simpl             = calcular_irrf_tabela(base_simpl, tabela)
    if (ir_simpl < ir_geral) or (ir_simpl == ir_geral and base_simpl <= base_geral):
        return ir_simpl, base_simpl
    return ir_geral, base_geral


def calcular_irrf_base60_mais_vantajoso_2026(bruto, inss, dependentes, ded_simpl, rt):
    if bruto is None or bruto <= 0:
        return 0.0, 0.0
    base60     = truncar(bruto * 0.60)
    dep_int    = 0 if (dependentes is None or pd.isna(dependentes)) else int(dependentes)
    red_dep    = truncar(dep_int * VALOR_DEP)
    base_legal = truncar(base60 - inss - red_dep)
    ir_legal   = calcular_irrf_2026_por_base(base_legal, rt)
    base_simpl = truncar(base60 - ded_simpl)
    ir_simpl   = calcular_irrf_2026_por_base(base_simpl, rt)
    if (ir_simpl < ir_legal) or (ir_simpl == ir_legal and base_simpl <= base_legal):
        return ir_simpl, base_simpl
    return ir_legal, base_legal


def calcular_irrf_base10_mais_vantajoso_2026(bruto, inss, dependentes, ded_simpl, rt):
    if bruto is None or bruto <= 0:
        return 0.0, 0.0
    base10     = truncar(bruto * 0.10)
    dep_int    = 0 if (dependentes is None or pd.isna(dependentes)) else int(dependentes)
    red_dep    = truncar(dep_int * VALOR_DEP)
    base_legal = truncar(base10 - inss - red_dep)
    ir_legal   = calcular_irrf_2026_por_base(base_legal, rt)
    base_simpl = truncar(base10 - ded_simpl)
    ir_simpl   = calcular_irrf_2026_por_base(base_simpl, rt)
    if (ir_simpl < ir_legal) or (ir_simpl == ir_legal and base_simpl <= base_legal):
        return ir_simpl, base_simpl
    return ir_legal, base_legal


def calcular_irrf_mais_vantajoso_2026_base100(base_bruta, dependentes, rt, ded_simpl):
    if base_bruta is None or base_bruta <= 0:
        return 0.0, 0.0, "nenhum"
    dep_int    = 0 if (dependentes is None or pd.isna(dependentes)) else int(dependentes)
    red_dep    = truncar(dep_int * VALOR_DEP)
    base_legal = truncar(base_bruta - red_dep)
    ir_legal   = calcular_irrf_2026_por_base(base_legal, rt)
    base_simpl = truncar(base_bruta - ded_simpl)
    ir_simpl   = calcular_irrf_2026_por_base(base_simpl, rt)
    if (ir_simpl < ir_legal) or (ir_simpl == ir_legal and base_simpl <= base_legal):
        return ir_simpl, base_simpl, "simplificada"
    return ir_legal, base_legal, "legal"


def calcular_irrf_acumulado_generico(rt_acum, inss_ded, deps, ano, tabela, ded_simpl):
    if rt_acum is None or rt_acum <= 0:
        return 0.0, 0.0
    dep_int  = max(0, 0 if (deps is None or pd.isna(deps)) else int(deps))
    red_dep  = truncar(dep_int * VALOR_DEP)
    b_legal  = max(truncar(rt_acum - inss_ded - red_dep), 0.0)
    b_simpl  = max(truncar(rt_acum - ded_simpl), 0.0)
    if ano == 2026:
        ir_l = calcular_irrf_2026_por_base(b_legal, rt_acum)
        ir_s = calcular_irrf_2026_por_base(b_simpl, rt_acum)
    else:
        ir_l = calcular_irrf_tabela(b_legal, tabela)
        ir_s = calcular_irrf_tabela(b_simpl, tabela)
    if (ir_s < ir_l) or (ir_s == ir_l and b_simpl <= b_legal):
        return ir_s, b_simpl
    return ir_l, b_legal


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
        c0s = str(c0).strip()
        pfx = "RELAÇÃO DE RENDIMENTOS - RPA:"
        r   = c0s[len(pfx):].strip() if c0s.startswith(pfx) else c0s
        if r.startswith("Empresa"):        codigo_empresa = df.iloc[i, 1]
        elif r.startswith("Razão Social"): razao_social   = df.iloc[i, 1]
        elif r.startswith("CNPJ"):         cnpj           = df.iloc[i, 1]
        elif r.startswith("Competencia"):  competencia    = df.iloc[i, 1]

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

    def _n(v):
        if v is None: return 0.0
        try:
            if pd.isna(v): return 0.0
        except Exception: pass
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return float(v)
        s = re.sub(r"[^0-9,\.\-]", "", str(v).strip())
        if s in ("", "-", ",", ".", "-.", "-,"): return 0.0
        if "." in s and "," in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")
        elif s.count(".") > 1:
            p = s.split(".")
            s = "".join(p[:-1]) + "." + p[-1]
        try: return float(s)
        except Exception: return 0.0

    registros = []
    for i in range(inicio, len(df)):
        linha = df.iloc[i]
        cod = linha[0] if len(linha) > 0 else None
        if cod is None or pd.isna(cod):
            continue
        try:
            if tem_cpf:
                nome=linha[1]; dep=linha[3]; eso=linha[4]; rpa=linha[5]
                atv=linha[6]; bruto=linha[7]; dpg=linha[8]; pen=linha[9]
                odesc=linha[10]; oprov=linha[11]; piss=linha[12]; diss=linha[13]
            else:
                nome=linha[1]; dep=linha[2]; eso=linha[3]; rpa=linha[4]
                atv=linha[5]; bruto=linha[6]; dpg=linha[7]; pen=linha[8]
                odesc=linha[9]; oprov=linha[10]; piss=linha[11]; diss=linha[12]

            if bruto is None or pd.isna(bruto):
                log.append(f"Aviso: linha {i+1} sem BRUTO. Código: {cod}. Pulando.")
                continue

            registros.append({
                "cod_contrib": cod, "nome": nome, "dependentes": dep,
                "esocial": eso, "rpa_num": rpa, "atividade": atv,
                "bruto": _n(bruto), "data_pagto": dpg,
                "pensao_alim": _n(pen), "outros_desc": _n(odesc),
                "outros_prov": _n(oprov), "perc_iss": _n(piss),
                "valor_iss": 0.0, "data_iss": diss, "linha_excel": i + 1,
            })
        except Exception as e:
            log.append(f"ERRO ao ler linha {i+1}: {e}")

    return {
        "codigo_empresa": codigo_empresa, "razao_social": razao_social,
        "cnpj": str(cnpj), "competencia": competencia, "registros": registros,
    }


# ==============================
# MONTAGEM DO REGISTRO (266 chars)
# ==============================
def montar_registro_lancamento(meta, reg, log, acum_mes):
    cod_emp  = meta["codigo_empresa"]
    comp_str = competencia_aaaamm(meta["competencia"])

    dt_pag    = (excel_date_to_datetime(reg.get("data_pagto")) or ultimo_dia_competencia(meta["competencia"]))
    str_pag   = "00000000" if dt_pag is None else dt_pag.strftime("%Y%m%d")
    ano_ir    = dt_pag.year if dt_pag is not None else None
    tabela_ir = tabela_ir_por_data_pagto(dt_pag)
    ded_simpl = deducao_simplificada_por_data_pagto_ou_ano(dt_pag)

    cod     = reg["cod_contrib"]
    deps    = reg["dependentes"]
    rpa_num = reg["rpa_num"]
    atv     = reg["atividade"]
    bruto   = limpar_negativo(reg["bruto"])
    perc_iss = limpar_negativo(reg.get("perc_iss", 0.0))
    pensao   = limpar_negativo(reg.get("pensao_alim", 0.0))
    odesc    = limpar_negativo(reg.get("outros_desc", 0.0))
    oprov    = limpar_negativo(reg.get("outros_prov", 0.0))

    dt_iss  = excel_date_to_datetime(reg.get("data_iss"))
    str_iss = "00000000" if dt_iss is None else dt_iss.strftime("%Y%m%d")

    esocial = reg.get("esocial")
    try:
        eso_int = int(esocial) if not pd.isna(esocial) else None
    except Exception:
        eso_int = None

    chave = chave_acumulacao_mes(meta, reg, dt_pag)
    if chave not in acum_mes:
        acum_mes[chave] = {
            "base_inss_empresa": 0.0, "inss_retido_empresa": 0.0,
            "outras_fontes_base": 0.0, "rend_trib_irrf": 0.0,
            "inss_dedutivel_irrf": 0.0, "irrf_retido": 0.0, "dependentes": 0,
        }
    ac = acum_mes[chave]

    sest = senat = 0.0
    base_orig = bruto
    aliq_inss = 0.11

    if eso_int in (712, 734):
        base_orig = truncar(bruto * 0.20)
        aliq_inss = 0.20 if eso_int == 734 else 0.11
        sest      = truncar(base_orig * 0.015)
        senat     = truncar(base_orig * 0.010)

    teto  = teto_inss_por_data_pagto(dt_pag)
    saldo = max(truncar(teto - max(truncar(ac.get("outras_fontes_base", 0.0)), 0.0)), 0.0)
    b_ant = truncar(ac["base_inss_empresa"])
    b_nov = truncar(b_ant + base_orig)
    b_lim = max(truncar(min(b_nov, saldo) - min(b_ant, saldo)), 0.0)
    inss  = max(truncar(b_lim * aliq_inss), 0.0)

    ac["base_inss_empresa"]   = b_nov
    ac["inss_retido_empresa"] = truncar(ac["inss_retido_empresa"] + inss)
    base_inss = b_lim

    rt_reg  = obter_rendimento_tributavel_irrf(bruto, eso_int)
    dep_out = max(0, 0 if (deps is None or pd.isna(deps)) else int(deps))
    deduz   = eso_int in (711, 712)

    ac["rend_trib_irrf"]      = truncar(ac["rend_trib_irrf"] + rt_reg)
    ac["inss_dedutivel_irrf"] = truncar(ac["inss_dedutivel_irrf"] + inss)
    ac["dependentes"]         = max(ac["dependentes"], dep_out)

    inss_ded  = ac["inss_dedutivel_irrf"] if deduz else 0.0
    ano_calc  = ano_ir if ano_ir in (2025, 2026) else 2025
    if ano_ir not in (2025, 2026):
        log.append(f"Aviso: ano desconhecido ({ano_ir}) para contrib {cod}; usando 2025.")

    ir_total, base_irrf = calcular_irrf_acumulado_generico(
        ac["rend_trib_irrf"], inss_ded, ac["dependentes"], ano_calc, tabela_ir, ded_simpl
    )

    ir_calc = max(truncar(ir_total - truncar(ac["irrf_retido"])), 0.0)
    ac["irrf_retido"] = truncar(ac["irrf_retido"] + ir_calc)

    if perc_iss and float(perc_iss) != 0.0:
        valor_iss = truncar(bruto * (perc_iss / 100.0))
    else:
        perc_iss, valor_iss = 0.0, 0.0

    valor_iss = limpar_negativo(valor_iss)
    base_inss = limpar_negativo(base_inss)
    sest      = limpar_negativo(sest)
    senat     = limpar_negativo(senat)
    inss      = limpar_negativo(inss)
    base_irrf = limpar_negativo(base_irrf)
    ir_calc   = limpar_negativo(ir_calc)

    try:
        registro = (
            fmt_int(cod_emp, 7)    + fmt_int(cod, 10)       + comp_str            +
            fmt_str(atv, 100)      + fmt_int(rpa_num, 10)   + fmt_num(bruto, 11)  +
            fmt_num(perc_iss, 5)   + fmt_num(valor_iss, 11) + str_iss             +
            fmt_num(base_inss, 11) + fmt_num(sest, 8)       + fmt_num(senat, 8)   +
            fmt_num(inss, 8)       + fmt_num(pensao, 11)    + fmt_num(odesc, 11)  +
            fmt_num(oprov, 11)     + str_pag                + fmt_num(base_irrf, 11) +
            fmt_int(dep_out, 3)    + fmt_num(ir_calc, 8)
        )
    except Exception as e:
        log.append(f"ERRO ao montar registro do contrib {cod}: {e}")
        return None

    if len(registro) != 266:
        log.append(f"ERRO: tamanho {len(registro)} != 266. Cód={cod_emp}, contrib={cod}")
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

        linhas, acum = [], {}
        for reg in meta["registros"]:
            l = montar_registro_lancamento(meta, reg, log, acum)
            if l is not None:
                linhas.append(l)

        if any(str(x).startswith("ERRO") for x in log):
            log.append("ERRO: Geração cancelada. TXT NÃO foi gerado.")
            return None
        if not linhas:
            log.append("ERRO: Nenhum registro válido foi gerado para o TXT.")
            return None

        log.append(f"Arquivo TXT gerado com {len(linhas)} registros.")
        return linhas

    except Exception:
        log.append("ERRO FATAL durante a geração do arquivo.")
        log.append(traceback.format_exc())
        return None


# ==============================
# UTILITÁRIO BGR
# ==============================
def carregar_bgr_bytes():
    caminho = os.path.join(os.path.dirname(os.path.abspath(__file__)), "bgr_base64.txt")
    if not os.path.exists(caminho):
        return None, "Arquivo bgr_base64.txt não encontrado na pasta do app."
    try:
        with open(caminho, "r", encoding="utf-8") as f:
            c = f.read().strip()
        if c.startswith("data:") and "," in c:
            c = c.split(",", 1)[1]
        return base64.b64decode(c), None
    except Exception as e:
        return None, f"Erro ao decodificar bgr_base64.txt: {e}"


# ==============================
# CSS GLOBAL — Tema TR / Domínio Escuro
# ==============================
TR_DARK_CSS = """
<style>
[data-testid="stAppViewContainer"],
[data-testid="stMain"], .main        { background-color: #0D1117 !important; }
[data-testid="stSidebar"]            { background-color: #0D1117 !important; border-right: 1px solid #21262D !important; }
[data-testid="stSidebarContent"]     { padding: 0 !important; }
html, body, [class*="css"]           { font-family: 'Segoe UI','Inter',sans-serif !important; color: #E6EDF3 !important; }
h1,h2,h3,h4,h5,h6                   { color: #E6EDF3 !important; font-weight: 700 !important; }

[data-testid="stButton"] button[kind="primary"] {
    background: linear-gradient(135deg,#FF6200,#E05500) !important;
    color:#fff !important; border:none !important; border-radius:6px !important;
    font-weight:700 !important; box-shadow:0 2px 10px rgba(255,98,0,.4) !important;
    transition:opacity .2s,transform .1s !important;
}
[data-testid="stButton"] button[kind="primary"]:hover   { opacity:.88 !important; transform:translateY(-1px) !important; }
[data-testid="stButton"] button[kind="primary"]:active  { transform:translateY(0) !important; }
[data-testid="stButton"] button[kind="primary"]:disabled { background:#2D333B !important; color:#484F58 !important; box-shadow:none !important; }

[data-testid="stButton"] button[kind="secondary"] {
    background-color:#21262D !important; color:#E6EDF3 !important;
    border:1px solid #30363D !important; border-radius:6px !important;
    font-weight:500 !important; transition:background-color .2s !important;
}
[data-testid="stButton"] button[kind="secondary"]:hover { background-color:#2D333B !important; border-color:#484F58 !important; }

[data-testid="stDownloadButton"] button {
    background:linear-gradient(135deg,#006AFF,#0054CC) !important;
    color:#fff !important; border:none !important; border-radius:6px !important;
    font-weight:700 !important; box-shadow:0 2px 10px rgba(0,106,255,.4) !important;
    transition:opacity .2s !important;
}
[data-testid="stDownloadButton"] button:hover { opacity:.88 !important; }

[data-testid="stFileUploader"] {
    background-color:#161B22 !important; border:1.5px dashed #30363D !important;
    border-radius:8px !important; transition:border-color .2s !important;
}
[data-testid="stFileUploader"]:hover { border-color:#FF6200 !important; }

hr { border-color:#21262D !important; }

::-webkit-scrollbar       { width:6px; height:6px; }
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

        # ── Cabeçalho ─────────────────────────────────────
        st.markdown("<div style='background:linear-gradient(180deg,#1C2128,#161B22);border-bottom:1px solid #21262D;padding:18px 16px 14px;'><div style='font-size:9px;font-weight:800;letter-spacing:.18em;color:#FF6200;text-transform:uppercase;margin-bottom:3px;'>Thomson Reuters</div><div style='font-size:15px;font-weight:700;color:#E6EDF3;'>Dom&#237;nio Sistemas</div><div style='font-size:11px;color:#8B949E;margin-top:2px;'>Gerador RPA &nbsp;&middot;&nbsp; Utilit&#225;rios</div></div>", unsafe_allow_html=True)

        st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)

        # ── Seção BGR: título ──────────────────────────────
        st.markdown("<div style='padding:0 14px;'><div style='border-left:3px solid #006AFF;padding-left:10px;margin-bottom:6px;'><span style='font-size:13px;font-weight:700;color:#E6EDF3;'>&#128444;&#65039; Arquivo BGR</span><br><span style='font-size:11px;color:#8B949E;'>Background &middot; Dom&#237;nio Sistemas</span></div></div>", unsafe_allow_html=True)

        st.caption("Baixe o arquivo BGR e importe no módulo Folha do Domínio Sistemas.")

        # ── Botão download BGR ─────────────────────────────
        bgr_bytes, bgr_erro = carregar_bgr_bytes()

        if bgr_bytes is not None:
            st.download_button(
                label="⬇  Baixar BGR para Domínio",
                data=bgr_bytes,
                file_name=NOME_ARQUIVO_BGR,
                mime="application/octet-stream",
                use_container_width=True,
                help=f"Salva como: {NOME_ARQUIVO_BGR}",
            )
            st.markdown(f"<div style='font-size:11px;color:#3FB950;margin:4px 2px 0;'>&#10004; Dispon&#237;vel &middot; {len(bgr_bytes):,} bytes</div>", unsafe_allow_html=True)
        else:
            st.markdown(f"<div style='background:#1A0D0D;border:1px solid #DA3633;border-radius:6px;padding:10px 12px;font-size:12px;color:#F85149;'>&#9888; BGR n&#227;o encontrado<br><span style='color:#8B949E;font-size:11px;'>{bgr_erro}</span></div>", unsafe_allow_html=True)

        st.divider()

        # ── Instruções de importação BGR ───────────────────
        st.markdown("<div style='font-size:10px;font-weight:700;letter-spacing:.1em;color:#8B949E;text-transform:uppercase;margin-bottom:10px;padding:0 2px;'>&#128196; Instru&#231;&#245;es de Importa&#231;&#227;o</div>", unsafe_allow_html=True)

        # Módulo — badge
        st.markdown("<div style='background:#1C2128;border:1px solid #21262D;border-radius:6px;padding:8px 12px;margin-bottom:10px;display:flex;align-items:center;gap:8px;'><span style='font-size:14px;'>&#128193;</span><div><div style='font-size:11px;font-weight:700;color:#FF6200;text-transform:uppercase;letter-spacing:.08em;'>M&#243;dulo Folha</div><div style='font-size:11px;color:#8B949E;margin-top:1px;'>Usu&#225;rio: <strong style='color:#C9D1D9;'>Gerente</strong></div></div></div>", unsafe_allow_html=True)

        # Passos oficiais — cada um em st.markdown flat de 1 linha
        passos_bgr = [
            ("1", "#FF6200", "No m&#243;dulo <strong>Folha</strong> com usu&#225;rio <strong>Gerente</strong>"),
            ("2", "#FF6200", "Acesse <strong>Relat&#243;rios &gt; Gerenciador de Relat&#243;rios</strong>"),
            ("3", "#FF6200", "Clique em <strong>[Novo]</strong> e depois em <strong>[Importar]</strong>"),
            ("4", "#FF6200", "No <strong>Dom&#237;nio Gerador</strong>, acesse o menu <strong>Utilit&#225;rios</strong>"),
            ("5", "#FF6200", "Clique em <strong>[...]</strong>, localize o BGR baixado e clique em <strong>[Abrir]</strong>"),
            ("6", "#FF6200", "Confira <strong>Classifica&#231;&#227;o</strong> e <strong>T&#237;tulo</strong> e clique em <strong>[Ok]</strong>"),
            ("7", "#FF6200", "Feche o <strong>Gerenciador de Relat&#243;rios</strong> para atualizar"),
            ("8", "#FF6200", "Reabra <strong>Relat&#243;rios &gt; Gerenciador de Relat&#243;rios</strong>"),
            ("9", "#FF6200", "Localize a pasta e selecione o arquivo importado"),
            ("10", "#D29922", "Informe os argumentos e clique em <strong>[Executar...]</strong>"),
        ]

        for num, cor, texto in passos_bgr:
            st.markdown(f"<div style='display:flex;align-items:flex-start;gap:8px;margin-bottom:7px;padding:0 2px;'><div style='min-width:20px;height:20px;background:{cor};border-radius:50%;font-size:10px;font-weight:800;color:#fff;display:flex;align-items:center;justify-content:center;flex-shrink:0;margin-top:1px;'>{num}</div><span style='font-size:11.5px;color:#C9D1D9;line-height:1.5;'>{texto}</span></div>", unsafe_allow_html=True)

        # Nota final de sucesso
        st.markdown("<div style='background:#0D2818;border:1px solid #238636;border-radius:6px;padding:8px 12px;margin-top:6px;font-size:11px;color:#3FB950;'>&#10003; Ap&#243;s importar, o relat&#243;rio estar&#225; dispon&#237;vel no Gerenciador de Relat&#243;rios.</div>", unsafe_allow_html=True)

        st.divider()

        # ── Rodapé ─────────────────────────────────────────
        st.markdown(f"<div style='text-align:center;font-size:11px;color:#484F58;line-height:1.7;padding-bottom:8px;'>Gerador RPA &middot; <span style='color:#FF6200;font-weight:700;'>{VERSAO}</span><br>Thomson Reuters &middot; Dom&#237;nio Sistemas</div>", unsafe_allow_html=True)


# ==============================
# MAIN
# ==============================
def main():
    st.set_page_config(
        page_title=f"Gerador RPA | {VERSAO}",
        page_icon="🧾",
        layout="wide",
    )

    st.markdown(TR_DARK_CSS, unsafe_allow_html=True)
    render_sidebar()

    # ── Header ────────────────────────────────────────────
    st.markdown(f"<div style='background:linear-gradient(135deg,#161B22,#1C2128);border:1px solid #21262D;border-top:3px solid #FF6200;border-radius:10px;padding:24px 28px 16px;margin-bottom:22px;'><div style='display:flex;align-items:center;gap:12px;'><span style='font-size:28px;'>&#129534;</span><div style='flex:1;'><div style='font-size:10px;font-weight:800;letter-spacing:.14em;color:#FF6200;text-transform:uppercase;margin-bottom:2px;'>Thomson Reuters &middot; Dom&#237;nio Sistemas</div><div style='font-size:20px;font-weight:700;color:#E6EDF3;'>Gerador de Arquivo TXT &#8212; RPA</div></div><div style='background:#21262D;border:1px solid #30363D;border-radius:20px;padding:3px 14px;font-size:12px;font-weight:700;color:#8B949E;'>{VERSAO}</div></div><p style='margin:8px 0 0 40px;font-size:13px;color:#8B949E;'>Importe o Excel e clique em <strong style='color:#E6EDF3;'>Gerar arquivo TXT</strong> para processar os lan&#231;amentos RPA.</p></div>", unsafe_allow_html=True)

    # ── Session state ──────────────────────────────────────
    if "log"        not in st.session_state: st.session_state.log        = [f"&#10004; Pronto &#8212; {VERSAO}"]
    if "txt_gerado" not in st.session_state: st.session_state.txt_gerado = None
    if "meta_info"  not in st.session_state: st.session_state.meta_info  = None

    # ── Upload ─────────────────────────────────────────────
    st.markdown("<div style='background:#161B22;border:1px solid #21262D;border-radius:8px;padding:14px 16px 2px;margin-bottom:14px;'><div style='font-size:10px;font-weight:700;letter-spacing:.1em;color:#8B949E;text-transform:uppercase;margin-bottom:8px;'>&#128194; Arquivo de entrada</div>", unsafe_allow_html=True)
    arquivo = st.file_uploader("Excel de entrada", type=["xlsx", "xls"], help="Planilha de Relação de Rendimentos — RPA", label_visibility="collapsed")
    st.markdown("</div>", unsafe_allow_html=True)

    # ── Botões ─────────────────────────────────────────────
    c1, c2 = st.columns([4, 1])
    with c1:
        gerar  = st.button("▶  Gerar arquivo TXT", disabled=(arquivo is None), use_container_width=True, type="primary")
    with c2:
        limpar = st.button("🗑  Limpar", use_container_width=True, type="secondary")

    if limpar:
        st.session_state.log        = ["&#128465; Limpo. Pronto para nova opera&#231;&#227;o."]
        st.session_state.txt_gerado = None
        st.session_state.meta_info  = None
        st.rerun()

    # ── Processamento ──────────────────────────────────────
    if gerar and arquivo is not None:
        st.session_state.log        = ["&#9203; Iniciando processamento..."]
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
                        "empresa":     _meta.get("razao_social", "—"),
                        "cnpj":        _meta.get("cnpj", "—"),
                        "competencia": competencia_aaaamm(_meta.get("competencia")),
                        "registros":   len(linhas),
                    }
            except Exception:
                pass
        st.rerun()

    # ── Cards de resultado ─────────────────────────────────
    if st.session_state.txt_gerado is not None:
        if st.session_state.meta_info:
            m    = st.session_state.meta_info
            comp = (f"{m['competencia'][4:6]}/{m['competencia'][:4]}" if len(m["competencia"]) == 6 else m["competencia"])

            c_emp, c_cnpj, c_comp, c_reg = st.columns(4)
            for col, label, valor, cor in [
                (c_emp,  "Empresa",     m["empresa"],   "#FF6200"),
                (c_cnpj, "CNPJ",        m["cnpj"],      "#006AFF"),
                (c_comp, "Competência", comp,           "#D29922"),
                (c_reg,  "Registros",   m["registros"], "#3FB950"),
            ]:
                fsize  = "20px" if label == "Registros" else "13px"
                fcolor = cor    if label == "Registros" else "#E6EDF3"
                with col:
                    st.markdown(f"<div style='background:#161B22;border:1px solid #21262D;border-top:2px solid {cor};border-radius:8px;padding:14px;margin-bottom:12px;'><div style='font-size:10px;font-weight:700;letter-spacing:.1em;color:#8B949E;text-transform:uppercase;margin-bottom:4px;'>{label}</div><div style='font-size:{fsize};font-weight:700;color:{fcolor};overflow:hidden;text-overflow:ellipsis;white-space:nowrap;' title='{valor}'>{valor}</div></div>", unsafe_allow_html=True)

        st.markdown("<div style='background:#0D2818;border:1px solid #238636;border-left:4px solid #3FB950;border-radius:8px;padding:12px 18px;margin-bottom:12px;display:flex;align-items:center;gap:10px;'><span style='font-size:18px;'>&#9989;</span><span style='font-size:14px;font-weight:600;color:#3FB950;'>Arquivo gerado com sucesso! Clique abaixo para baixar.</span></div>", unsafe_allow_html=True)

        st.download_button(
            label="⬇  Baixar arquivo TXT",
            data=st.session_state.txt_gerado,
            file_name=NOME_ARQUIVO_TXT,
            mime="text/plain",
            use_container_width=True,
            type="primary",
        )

    # ── Log ────────────────────────────────────────────────
    st.markdown("<div style='margin-top:22px;'></div>", unsafe_allow_html=True)

    tem_erro  = any(str(l).startswith("ERRO")  for l in st.session_state.log)
    tem_aviso = any(str(l).startswith("Aviso") for l in st.session_state.log)

    if tem_erro:
        acc, brd, bg, ico, lbl = "#F85149","#DA3633","#1A0D0D","&#128308;","Erros detectados"
    elif tem_aviso:
        acc, brd, bg, ico, lbl = "#D29922","#9E6A03","#1A1500","&#128993;","Avisos"
    else:
        acc, brd, bg, ico, lbl = "#3FB950","#238636","#0D1A10","&#128994;","OK"

    log_txt = "\n".join(str(l) for l in st.session_state.log)

    st.markdown(f"<div style='background:#161B22;border:1px solid #21262D;border-top:2px solid {acc};border-radius:8px;overflow:hidden;'><div style='background:#1C2128;padding:10px 16px;border-bottom:1px solid #21262D;display:flex;align-items:center;gap:8px;'><span>{ico}</span><span style='font-size:11px;font-weight:700;letter-spacing:.08em;color:#8B949E;text-transform:uppercase;'>Log de processamento &#8212; {lbl}</span></div><div style='background:{bg};border-left:3px solid {brd};margin:12px;border-radius:4px;padding:12px 14px;font-family:Cascadia Code,Fira Code,Consolas,monospace;font-size:12.5px;line-height:1.7;white-space:pre-wrap;max-height:280px;overflow-y:auto;color:#C9D1D9;'>{log_txt}</div></div>", unsafe_allow_html=True)

    # ── Rodapé ─────────────────────────────────────────────
    st.markdown(f"<div style='margin-top:28px;padding-top:12px;border-top:1px solid #21262D;display:flex;justify-content:space-between;align-items:center;'><span style='font-size:11px;color:#484F58;'>Thomson Reuters &middot; Dom&#237;nio Sistemas &middot; Gerador RPA</span><span style='font-size:11px;font-weight:800;letter-spacing:.08em;color:#FF6200;'>{VERSAO}</span></div>", unsafe_allow_html=True)


if __name__ == "__main__":
    main()
