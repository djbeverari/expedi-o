import re
import pandas as pd
import sys
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from datetime import datetime
import sys, io
if hasattr(sys.stdout, "buffer"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

ARQ_EXTRATO        = "ESTOQUE_AJUSTADO.xlsx"
ARQ_ESTOQUE_MAXIMO = "Estoque_Máximo.xlsx"
ARQ_ESTOQUE_MINIMO = "Estoque_Mínimo.xlsx"
ARQ_ESTOQUE_IDEAL  = "Estoque_Ideal.xlsx"
ARQ_COMPRAS        = "COMPRAS_AJUSTADO.xlsx"
FILIAL_MATRIZ      = "SANTA RITA"

IMPRESSAO_COLORIDA = False

# ═══════════════════════════════════════════════════════════════════════════
# PARÂMETROS CONFIGURÁVEIS - EDITE AQUI
# ═══════════════════════════════════════════════════════════════════════════

FATOR_VENDAS_MAX = 1.5
FATOR_VENDAS_MIN = 2.0

FATOR_TAMANHO_ALTO = 2.0
FATOR_TAMANHO_BAIXO = 3.0

FATOR_ML_TRAD_MAX = 1.5
FATOR_ML_TRAD_MIN = 2.0

FATOR_MC_FIT_MAX = 1.5
FATOR_MC_FIT_MIN = 2.0

TETO_REFERENCIA = 5000

# ═══════════════════════════════════════════════════════════════════════════

TAM_ML_FIT = ['01','02','03','04','05','06','07','08']
TAM_ML_TRADICIONAL = ['02','03','04','05','06','07']
TAM_MC_ADULTO = ['02','03','04','05','06','07']
TAM_JUVENIL   = ['02','04','06','08','10','12','14','16']

LIMITES_PRODUTO_COR = {
    ('05.01.03', '1'): {'min': 100, 'max': 300},
    ('05.01.68', '1'): {'min': 80, 'max': 200},
}

# Limite máximo padrão aplicado a todos os produtos/cores não listados em LIMITES_PRODUTO_COR
# (exceto produtos SOMENTE ENCOMENDAS e os listados explicitamente em LIMITES_PRODUTO_COR)
LIMITE_MAX_PADRAO = 200

# Produtos especiais que sempre mostram "SOMENTE ENCOMENDAS"
PRODUTOS_SOMENTE_ENCOMENDAS = ['050198', '050233']  # 05.01.98 e 05.02.33

def verificar_produto_especial(produto_formatado, produto_raw, df_extrato):
    """
    CORREÇÃO 3: Produtos especiais (05.01.98 e 05.02.33) 
    SEMPRE mostram "SOMENTE ENCOMENDAS".
    Não geram alertas normais.
    """
    produto_limpo = produto_formatado.replace('.', '')
    if produto_limpo in PRODUTOS_SOMENTE_ENCOMENDAS:
        # Produto especial: SEMPRE retorna "SOMENTE ENCOMENDAS"
        return ["SOMENTE ENCOMENDAS"], True
    return [], False  # Não é produto especial



def formatar_codigo_produto(codigo):
    """Retorna código formatado. Se já tem pontos, retorna direto."""
    s = str(codigo)
    if '.' in s:
        return s
    if len(s) >= 5:
        return f"05.{s[1:3]}.{s[3:]}"
    else:
        s = s.zfill(6)
        return f"{s[0:2]}.{s[2:4]}.{s[4:]}"


def formatar_numero(valor):
    """Padrão brasileiro: 1.234"""
    return f"{valor:,.0f}".replace(',', '.')


# Reduzir descrição para 8 caracteres para garantir que não vaze
def limpar_descricao(desc, max_chars=16):
    """Remove 'CAMISA' e 'LINEA' e abrevia se necessário."""
    if not desc:
        return ""
    desc = str(desc).upper()
    desc = desc.replace('CAMISA', '').replace('LINEA', '').strip()
    if len(desc) > max_chars:
        desc = desc[:max_chars-2] + '..'
    return desc


def verificar_limites_produto(produto_formatado, cor, total_loja):
    chave = (produto_formatado, cor)
    if chave in LIMITES_PRODUTO_COR:
        # Produto com limites específicos cadastrados
        limites = LIMITES_PRODUTO_COR[chave]
        if total_loja < limites['min']:
            return [f"! ABAIXO MÍN.({limites['min']})"]
        elif total_loja > limites['max']:
            return [f"! ACIMA MÁX.({limites['max']})"]
        return []
    else:
        # Regra geral: limite máximo padrão para todos os demais produtos
        if total_loja > LIMITE_MAX_PADRAO:
            return [f"! ACIMA MÁX.({LIMITE_MAX_PADRAO})"]
        return []


def carregar_dados():
    try:
        df_extrato = pd.read_excel(ARQ_EXTRATO)
        print(f"OK Extrato carregado: {len(df_extrato)} linhas")
        df_max = pd.read_excel(ARQ_ESTOQUE_MAXIMO)
        print(f"OK Estoque máximo: {len(df_max)} lojas")
        df_min = pd.read_excel(ARQ_ESTOQUE_MINIMO)
        print(f"OK Estoque mínimo: {len(df_min)} lojas")
        
        try:
            df_vendas = pd.read_excel(ARQ_ESTOQUE_IDEAL, sheet_name='estoque_ideal_percentual')
            print(f"OK Estoque Ideal (percentual): {len(df_vendas)} lojas")
        except FileNotFoundError:
            df_vendas = None
            print(f"  ({ARQ_ESTOQUE_IDEAL} não encontrado)")
        except Exception as e:
            df_vendas = None
            print(f"  (Erro ao carregar aba estoque_ideal_percentual: {e})")
        
        try:
            df_ideal = pd.read_excel(ARQ_ESTOQUE_IDEAL, sheet_name='estoque_ideal')
            print(f"OK Estoque Ideal (absoluto): {len(df_ideal)} lojas")
        except FileNotFoundError:
            df_ideal = None
            print(f"  ({ARQ_ESTOQUE_IDEAL} - aba estoque_ideal não encontrada)")
        except Exception as e:
            df_ideal = None
            print(f"  (Erro ao carregar aba estoque_ideal: {e})")
        
        try:
            df_compras = pd.read_excel(ARQ_COMPRAS)
            print(f"OK Compras: {len(df_compras)} registros")
        except FileNotFoundError:
            df_compras = None
            print(f"  ({ARQ_COMPRAS} não encontrado)")
        
        # Calcular fator de ajuste: soma ideais lojas / total rede (lojas + matriz)
        if df_ideal is not None:
            cols_num = ['MANGA  LONGA', 'MANGA CURTA', 'JUVENIL']
            soma_ideais = df_ideal[cols_num].sum().sum()
            cols_tam_ext = [c for c in df_extrato.columns if c.isdigit()]
            total_rede_geral = df_extrato[
                df_extrato['Tipo Produto'].isin(['ADULTO', 'ADULTO EXTRA', 'JUVENIL'])
            ][cols_tam_ext].sum().sum()
            fator_3b = soma_ideais / total_rede_geral if total_rede_geral > 0 else 1.0
            print(f"OK Fator de ajuste calculado: {soma_ideais:.0f} / {total_rede_geral:.0f} = {fator_3b:.4f}")
        else:
            fator_3b = 1.0

        return df_extrato, df_max, df_min, df_vendas, df_ideal, df_compras, fator_3b
    except FileNotFoundError as e:
        print(f"\nERRO Arquivo não encontrado: {e}")
        sys.exit(1)


def listar_lojas_disponiveis(df_extrato, df_max):
    return sorted(set(df_extrato['Filial'].unique()) & set(df_max['FILIAL'].unique()))


def calcular_estoque_loja(df_extrato, filial):
    df_loja = df_extrato[df_extrato['Filial'] == filial].copy()
    if df_loja.empty:
        return None
    cols_tam = [c for c in df_loja.columns if c.isdigit()]
    r = {'MANGA LONGA_ADULTO':0, 'MANGA CURTA_ADULTO':0, 'JUVENIL_TOTAL':0,
         'ML_FIT':0, 'ML_TRADICIONAL':0, 'ML_TRADICIONAL_EXTRA':0,
         'MC_FIT':0, 'MC_TRADICIONAL':0, 'MC_TRADICIONAL_EXTRA':0}
    for _, row in df_loja.iterrows():
        sg     = str(row['Subgrupo']).strip().upper()
        tipo   = str(row['Tipo Produto']).strip().upper()
        subcat = str(row['Subcategoria  Produto']).strip().upper()
        total  = row[cols_tam].sum()
        
        if tipo in ['ADULTO', 'ADULTO EXTRA']:
            if 'MANGA LONGA' in sg:
                r['MANGA LONGA_ADULTO'] += total
                if tipo == 'ADULTO EXTRA':
                    r['ML_TRADICIONAL_EXTRA'] += total
                elif subcat == 'SLIM':
                    r['ML_FIT'] += total
                elif subcat == 'TRADICIONAL':
                    r['ML_TRADICIONAL'] += total
            elif 'MANGA CURTA' in sg:
                r['MANGA CURTA_ADULTO'] += total
                if tipo == 'ADULTO EXTRA':
                    r['MC_TRADICIONAL_EXTRA'] += total
                elif subcat == 'SLIM':
                    r['MC_FIT'] += total
                elif subcat == 'TRADICIONAL':
                    r['MC_TRADICIONAL'] += total
        
        if tipo == 'JUVENIL':
            r['JUVENIL_TOTAL'] += total
    return r


def obter_compras_a_receber(df_compras, filial):
    """Retorna lista de alertas de compras a receber para a loja."""
    if df_compras is None or df_compras.empty:
        return {'ML': [], 'MC': [], 'JUV': []}
    
    df_compras_limpo = df_compras[
        df_compras['Filial'].str.strip().str.upper().str.startswith('LOJA')
    ].copy()
    
    filial_clean = filial.strip().upper()
    df_loja = df_compras_limpo[df_compras_limpo['Filial'].str.strip().str.upper() == filial_clean]
    
    alertas = {'ML': [], 'MC': [], 'JUV': []}
    
    for _, row in df_loja.iterrows():
        try:
            qtd = int(row['Qtde Entregar'])
            desc = str(row['Desc Produto']).strip()
            forn = str(row['Fornecedor']).strip()
            subgrupo = str(row['Subgrupo Produto']).strip().upper()
            
            msg = f"{qtd} {desc.upper()} À RECEBER DE {forn.upper()}"
            
            if 'MANGA LONGA' in subgrupo:
                alertas['ML'].append(msg)
            elif 'MANGA CURTA' in subgrupo:
                alertas['MC'].append(msg)
            elif 'JUVENIL' in subgrupo:
                alertas['JUV'].append(msg)
            else:
                alertas['ML'].append(msg)
        except (KeyError, ValueError) as e:
            print(f"  Aviso: Erro ao processar linha de compras: {e}")
            continue
    
    return alertas


def obter_limites(df_max, df_min, filial):
    row_max = df_max[df_max['FILIAL'] == filial]
    if row_max.empty:
        return None
    row_min = df_min[df_min['FILIAL'] == filial]
    maximo = {'MANGA LONGA': int(row_max['MANGA  LONGA'].values[0]),
              'MANGA CURTA': int(row_max['MANGA CURTA'].values[0]),
              'JUVENIL'    : int(row_max['JUVENIL'].values[0])}
    minimo = ({'MANGA LONGA': int(row_min['MANGA  LONGA'].values[0]),
               'MANGA CURTA': int(row_min['MANGA CURTA'].values[0]),
               'JUVENIL'    : int(row_min['JUVENIL'].values[0])}
              if not row_min.empty else {k: None for k in maximo})
    return maximo, minimo


def obter_ideal(df_ideal, filial):
    """Retorna dict com valores de estoque ideal por categoria para a filial."""
    if df_ideal is None:
        return None
    row = df_ideal[df_ideal['FILIAL'] == filial]
    if row.empty:
        return None
    return {
        'MANGA LONGA': int(row['MANGA  LONGA'].values[0]),
        'MANGA CURTA': int(row['MANGA CURTA'].values[0]),
        'JUVENIL':     int(row['JUVENIL'].values[0]),
    }


def agrupar_alertas_tamanhos(pouco, muito):
    """
    Converte listas de tamanhos em alertas agrupados e simplificados.
    Ex: pouco=[02,04,05], muito=[03] → ["POUCO TAM 02, 04 E 05", "MUITO TAM 03"]
    """
    def formatar_lista(tams):
        tams = sorted(tams, key=lambda x: int(x))
        if len(tams) == 1:
            return f"TAM {tams[0]}"
        return "TAM " + ", ".join(tams[:-1]) + f" E {tams[-1]}"

    alertas = []
    if pouco:
        alertas.append(f"POUCO {formatar_lista(pouco)}")
    if muito:
        alertas.append(f"MUITO {formatar_lista(muito)}")
    return alertas


def analisar_distribuicao_tamanhos(df_extrato, filial, produto_raw, cor, tam_cols, total_loja,
                                   subgrupo=None, tipo_produto=None, subcat=None):
    """Análise: % por tamanho vs SANTA RITA (matriz). Retorna alertas agrupados POUCO/MUITO."""
    if total_loja <= 15:
        return []

    df_prod_matriz = df_extrato[
        (df_extrato['Filial'] == FILIAL_MATRIZ) &
        (df_extrato['Produto'] == produto_raw)
    ]
    if df_prod_matriz.empty:
        return []

    totais_matriz = df_prod_matriz[tam_cols].sum()
    total_matriz = totais_matriz.sum()
    if total_matriz == 0:
        return []

    pct_matriz = {t: totais_matriz[t] / total_matriz * 100 for t in tam_cols}

    df_loja_prod = df_extrato[
        (df_extrato['Filial'] == filial) &
        (df_extrato['Produto'] == produto_raw) &
        (df_extrato['Cor / Variante'] == cor)
    ]
    if df_loja_prod.empty:
        return []

    totais_loja = df_loja_prod[tam_cols].sum()
    total = totais_loja.sum()
    if total == 0:
        return []

    tamanhos_excluir = ['01']

    # Tamanho 07 excluído para ML Adulto
    if (subgrupo and 'MANGA LONGA' in subgrupo.upper() and
            tipo_produto and tipo_produto.upper() == 'ADULTO'):
        tamanhos_excluir.append('07')

    # Tamanho 02 excluído para ML Tradicional
    if (subgrupo and 'MANGA LONGA' in subgrupo.upper() and
            subcat and subcat.upper() == 'TRADICIONAL'):
        tamanhos_excluir.append('02')

    pouco, muito = [], []
    for tam in tam_cols:
        if totais_matriz[tam] == 0 or pct_matriz[tam] == 0:
            continue
        if tam in tamanhos_excluir:
            continue

        pct_loja = (totais_loja[tam] / total * 100) if total > 0 else 0

        if totais_loja[tam] > 0 and pct_loja > (pct_matriz[tam] * FATOR_TAMANHO_ALTO):
            muito.append(tam)
        elif pct_loja < (pct_matriz[tam] / FATOR_TAMANHO_BAIXO):
            pouco.append(tam)

    return agrupar_alertas_tamanhos(pouco, muito)


def analisar_distribuicao_total_tabela(df_extrato, filial, tam_cols, totais_loja, categoria=""):
    """
    NOVA FUNÇÃO: Analisa a distribuição de tamanhos no TOTAL de cada tabela.
    Compara o total da loja vs SANTA RITA (matriz) por tamanho.
    
    Retorna lista de alertas agrupados tipo: "POUCO TAM 02, 04 E 05" ou "MUITO TAM 03"
    """
    total_loja = sum(totais_loja.values())
    
    if total_loja <= 15:
        return []  # Estoque muito baixo, não analisar
    
    # Buscar estoque da MATRIZ (Santa Rita) para a mesma categoria
    df_matriz = df_extrato[df_extrato['Filial'] == FILIAL_MATRIZ]
    
    if df_matriz.empty:
        return []
    
    totais_matriz = df_matriz[tam_cols].sum()
    total_matriz = totais_matriz.sum()
    
    if total_matriz == 0:
        return []
    
    # Calcular percentuais
    pct_matriz = {t: (totais_matriz[t] / total_matriz * 100) if total_matriz > 0 else 0 
                  for t in tam_cols}
    pct_loja = {t: (totais_loja[t] / total_loja * 100) if total_loja > 0 else 0 
                for t in tam_cols}
    
    pouco, muito = [], []
    for tam in tam_cols:
        if tam == '01':
            continue
        if totais_matriz[tam] == 0 or pct_matriz[tam] == 0:
            continue
        if totais_loja[tam] > 0 and pct_loja[tam] > (pct_matriz[tam] * FATOR_TAMANHO_ALTO):
            muito.append(tam)
        elif pct_loja[tam] < (pct_matriz[tam] / FATOR_TAMANHO_BAIXO):
            pouco.append(tam)

    return agrupar_alertas_tamanhos(pouco, muito)


def analisar_estoque_vs_vendas(df_extrato, df_vendas, filial, produto_raw,
                               categoria='MANGA  LONGA', fator_3b=1.0, cor=None):
    """Análise: estoque vs % estoque ideal, calculado por cor individualmente.
    Referência = total_rede_da_cor × fator_de_ajuste.
    Se essa referência ainda > TETO_REFERENCIA, usa o teto fixo.
    """
    if df_vendas is None:
        return []

    row_vendas = df_vendas[df_vendas['FILIAL'] == filial]
    if row_vendas.empty:
        return []

    if categoria not in row_vendas.columns:
        return []
    pct_vendas_loja = row_vendas[categoria].values[0]

    if pct_vendas_loja < 1:
        pct_vendas_loja = pct_vendas_loja * 100

    colunas_tam = [c for c in df_extrato.columns if c.isdigit()]

    # Filtrar rede e loja pela cor quando fornecida
    df_prod_rede = df_extrato[df_extrato['Produto'] == produto_raw]
    df_prod_loja = df_extrato[
        (df_extrato['Filial'] == filial) &
        (df_extrato['Produto'] == produto_raw)
    ]
    if cor is not None:
        df_prod_rede = df_prod_rede[df_prod_rede['Cor / Variante'] == cor]
        df_prod_loja = df_prod_loja[df_prod_loja['Cor / Variante'] == cor]

    total_rede = df_prod_rede[colunas_tam].sum().sum()
    if total_rede == 0:
        return []

    total_loja = df_prod_loja[colunas_tam].sum().sum()

    # Ajustar referência pelo peso da matriz na rede
    referencia = total_rede * fator_3b
    if referencia > TETO_REFERENCIA:
        referencia = TETO_REFERENCIA

    pct_estoque_loja = (total_loja / referencia * 100) if referencia > 0 else 0

    limite_min = pct_vendas_loja / FATOR_VENDAS_MIN
    limite_max = pct_vendas_loja * FATOR_VENDAS_MAX

    if pct_estoque_loja > limite_max and total_loja > 25:
        return ["EST. ACIMA ESPERADO"]
    elif pct_estoque_loja < limite_min:
        return ["EST. ABAIXO ESPERADO"]
    return []


def montar_tabela_subcat(df_extrato, df_vendas, filial, sg, tipos, subcat, tams, fator_3b=1.0):
    """Monta tabela com APENAS produtos disponíveis na matriz."""
    mask_mtz = ((df_extrato['Filial'] == FILIAL_MATRIZ) &
                (df_extrato['Subgrupo'].str.contains(sg, na=False)) &
                (df_extrato['Tipo Produto'].isin(tipos)) &
                (df_extrato['Subcategoria  Produto'] == subcat))
    d_mtz = df_extrato[mask_mtz].groupby('Produto')[tams].sum()
    d_mtz['T'] = d_mtz.sum(axis=1)
    sel = d_mtz[d_mtz['T'] > 0].index.tolist()
    
    mask_loja = ((df_extrato['Filial'] == filial) &
                 (df_extrato['Subgrupo'].str.contains(sg, na=False)) &
                 (df_extrato['Tipo Produto'].isin(tipos)) &
                 (df_extrato['Subcategoria  Produto'] == subcat))
    df_loja = df_extrato[mask_loja].copy()
    
    tot_por_prod = df_loja.groupby('Produto')[tams].sum()
    tot_por_prod['T'] = tot_por_prod.sum(axis=1)
    
    df_sel = df_loja[df_loja['Produto'].isin(sel)].copy()
    by_pc  = df_sel.groupby(['Produto', 'Cor / Variante'])[tams].sum().reset_index()
    by_pc['TOTAL'] = by_pc[tams].sum(axis=1)
    
    presentes = set(by_pc['Produto'].unique())
    ausentes  = [p for p in sel if p not in presentes]
    if ausentes:
        df_mtz_aus = df_extrato[
            (df_extrato['Filial'] == FILIAL_MATRIZ) &
            (df_extrato['Subgrupo'].str.contains(sg, na=False)) &
            (df_extrato['Tipo Produto'].isin(tipos)) &
            (df_extrato['Subcategoria  Produto'] == subcat) &
            (df_extrato['Produto'].isin(ausentes))
        ].groupby(['Produto', 'Cor / Variante'])[tams].sum().reset_index()
        df_mtz_aus[tams] = 0
        df_mtz_aus['TOTAL'] = 0
        by_pc = pd.concat([by_pc, df_mtz_aus], ignore_index=True)
    
    by_pc = by_pc.sort_values('TOTAL', ascending=False)
    mtz_set = set(sel)
    
    # CORREÇÃO: Mostrar tamanhos que existem na LOJA ou na MATRIZ
    df_mtz_tams = df_extrato[
        (df_extrato['Filial'] == FILIAL_MATRIZ) &
        (df_extrato['Subgrupo'].str.contains(sg, na=False)) &
        (df_extrato['Tipo Produto'].isin(tipos)) &
        (df_extrato['Subcategoria  Produto'] == subcat)
    ]
    tams_matriz = [t for t in tams if df_mtz_tams[t].sum() > 0]
    tams_loja = [t for t in tams if by_pc[t].sum() > 0]
    tam_ok = sorted(set(tams_matriz + tams_loja), key=lambda x: tams.index(x))
    
    tem_desc_col = 'Desc Produto' in df_extrato.columns
    
    linhas = []
    for _, row in by_pc.iterrows():
        prod_raw = row['Produto']
        prod = formatar_codigo_produto(prod_raw)
        cor = row['Cor / Variante']
        cor_str = str(cor)
        total_linha = int(row['TOTAL'])
        
        descricao = ""
        if tem_desc_col:
            df_prod_desc = df_extrato[df_extrato['Produto'] == prod_raw]
            if not df_prod_desc.empty:
                desc_val = df_prod_desc.iloc[0]['Desc Produto']
                if pd.notna(desc_val):
                    descricao = limpar_descricao(desc_val)
        
        df_prod_info = df_extrato[df_extrato['Produto'] == prod_raw]
        tipo_produto_linha = df_prod_info.iloc[0]['Tipo Produto'] if not df_prod_info.empty else None
        subgrupo_linha = df_prod_info.iloc[0]['Subgrupo'] if not df_prod_info.empty else None
        
        alertas = []
        
        # CORREÇÃO 3: Verificar se é produto especial PRIMEIRO
        alerta_especial, eh_especial = verificar_produto_especial(prod, prod_raw, df_extrato)
        
        if eh_especial:
            # Produto especial: só mostra "SOMENTE ENCOMENDAS", pula outros alertas
            alertas = alerta_especial
        else:
            # Produto normal: gera alertas normais
            alertas.extend(verificar_limites_produto(prod, cor_str, total_linha))
            alertas.extend(analisar_distribuicao_tamanhos(
                df_extrato, filial, prod_raw, cor, tam_ok, total_linha, subgrupo_linha, tipo_produto_linha, subcat=subcat))
            # Determina categoria para buscar percentual correto no Estoque_Ideal
            if 'MANGA LONGA' in sg.upper():
                cat_ideal = 'MANGA  LONGA'
            elif 'MANGA CURTA' in sg.upper():
                cat_ideal = 'MANGA CURTA'
            else:
                cat_ideal = 'JUVENIL'
            alertas.extend(analisar_estoque_vs_vendas(
                df_extrato, df_vendas, filial, prod_raw, cat_ideal, fator_3b=fator_3b, cor=cor))
        
        linhas.append({
            'produto': prod,
            'descricao': descricao,
            'cor': cor_str,
            'tamanhos': {t: int(row[t]) for t in tam_ok},
            'total': total_linha,
            'anotacao': " | ".join(alertas) if alertas else "",
        })
    
    nao_sel = set(tot_por_prod.index) - set(sel)
    df_outros = df_loja[df_loja['Produto'].isin(nao_sel)]
    s = df_outros[tams].sum()
    s_filt = {t: int(s[t]) for t in tam_ok}
    label_outros = "Outros Fit" if subcat == 'SLIM' else "Outros Tradicional"
    outros = {'label': label_outros, 'tamanhos': s_filt, 'total': int(s.sum())}
    
    totais_tam = df_loja[tams].sum()
    total_tam  = {t: int(totais_tam[t]) for t in tam_ok}
    
    # CORREÇÃO 4: Analisar distribuição de tamanhos no TOTAL da tabela
    alertas_total = analisar_distribuicao_total_tabela(df_extrato, filial, tam_ok, total_tam)
    
    return {
        'linhas': linhas, 
        'outros': outros, 
        'total_tam': total_tam, 
        'tam_cols': tam_ok,
        'alertas_total': alertas_total  # NOVO: Alertas do total da tabela
    }


def montar_tabela_juvenil(df_extrato, df_vendas, filial, tams, fator_3b=1.0):
    """Tabela juvenil com APENAS produtos disponíveis na matriz."""
    mask_mtz = (df_extrato['Filial'] == FILIAL_MATRIZ) & (df_extrato['Tipo Produto'] == 'JUVENIL')
    d_mtz = df_extrato[mask_mtz].groupby('Produto')[tams].sum()
    d_mtz['T'] = d_mtz.sum(axis=1)
    sel = d_mtz[d_mtz['T'] > 0].index.tolist()
    
    mask_loja = (df_extrato['Filial'] == filial) & (df_extrato['Tipo Produto'] == 'JUVENIL')
    df_loja = df_extrato[mask_loja].copy()
    tot_por_prod = df_loja.groupby('Produto')[tams].sum()
    tot_por_prod['T'] = tot_por_prod.sum(axis=1)
    
    df_sel = df_loja[df_loja['Produto'].isin(sel)].copy()
    by_pc  = df_sel.groupby(['Produto', 'Cor / Variante'])[tams].sum().reset_index()
    by_pc['TOTAL'] = by_pc[tams].sum(axis=1)
    
    presentes = set(by_pc['Produto'].unique())
    ausentes  = [p for p in sel if p not in presentes]
    if ausentes:
        df_mtz_aus = df_extrato[
            (df_extrato['Filial'] == FILIAL_MATRIZ) &
            (df_extrato['Tipo Produto'] == 'JUVENIL') &
            (df_extrato['Produto'].isin(ausentes))
        ].groupby(['Produto', 'Cor / Variante'])[tams].sum().reset_index()
        df_mtz_aus[tams] = 0
        df_mtz_aus['TOTAL'] = 0
        by_pc = pd.concat([by_pc, df_mtz_aus], ignore_index=True)
    
    by_pc = by_pc.sort_values('TOTAL', ascending=False)
    mtz_set = set(sel)
    
    # CORREÇÃO: Mostrar tamanhos que existem na LOJA ou na MATRIZ
    df_mtz_tams = df_extrato[
        (df_extrato['Filial'] == FILIAL_MATRIZ) &
        (df_extrato['Tipo Produto'] == 'JUVENIL')
    ]
    tams_matriz = [t for t in tams if df_mtz_tams[t].sum() > 0]
    tams_loja = [t for t in tams if by_pc[t].sum() > 0]
    tam_ok = sorted(set(tams_matriz + tams_loja), key=lambda x: tams.index(x))
    
    tem_desc_col = 'Desc Produto' in df_extrato.columns
    
    linhas = []
    for _, row in by_pc.iterrows():
        prod_raw = row['Produto']
        prod = formatar_codigo_produto(prod_raw)
        cor = row['Cor / Variante']
        cor_str = str(cor)
        total_linha = int(row['TOTAL'])
        
        descricao = ""
        if tem_desc_col:
            df_prod_desc = df_extrato[df_extrato['Produto'] == prod_raw]
            if not df_prod_desc.empty:
                desc_val = df_prod_desc.iloc[0]['Desc Produto']
                if pd.notna(desc_val):
                    descricao = limpar_descricao(desc_val)
        
        df_prod_info = df_extrato[df_extrato['Produto'] == prod_raw]
        tipo_produto_linha = df_prod_info.iloc[0]['Tipo Produto'] if not df_prod_info.empty else None
        subgrupo_linha = df_prod_info.iloc[0]['Subgrupo'] if not df_prod_info.empty else None
        
        alertas = []
        
        # CORREÇÃO 3: Verificar se é produto especial PRIMEIRO
        alerta_especial, eh_especial = verificar_produto_especial(prod, prod_raw, df_extrato)
        
        if eh_especial:
            # Produto especial: só mostra "SOMENTE ENCOMENDAS", pula outros alertas
            alertas = alerta_especial
        else:
            # Produto normal: gera alertas normais
            alertas.extend(verificar_limites_produto(prod, cor_str, total_linha))
            alertas.extend(analisar_distribuicao_tamanhos(
                df_extrato, filial, prod_raw, cor, tam_ok, total_linha, subgrupo_linha, tipo_produto_linha))
            alertas.extend(analisar_estoque_vs_vendas(
                df_extrato, df_vendas, filial, prod_raw, 'JUVENIL', fator_3b=fator_3b, cor=cor))
        
        linhas.append({
            'produto': prod,
            'descricao': descricao,
            'cor': cor_str,
            'tamanhos': {t: int(row[t]) for t in tam_ok},
            'total': total_linha,
            'anotacao': " | ".join(alertas) if alertas else "",
        })
    
    nao_sel = set(tot_por_prod.index) - set(sel)
    df_outros = df_loja[df_loja['Produto'].isin(nao_sel)]
    s = df_outros[tams].sum()
    s_filt = {t: int(s[t]) for t in tam_ok}
    outros = {'label': 'Outros', 'tamanhos': s_filt, 'total': int(s.sum())}
    
    totais_tam = df_loja[tams].sum()
    total_tam  = {t: int(totais_tam[t]) for t in tam_ok}
    
    return {'linhas': linhas, 'outros': outros, 'total_tam': total_tam, 'tam_cols': tam_ok}


def status_estoque(atual, maximo, minimo, ideal=None):
    msgs = []
    if atual > maximo:
        cor = colors.HexColor('#e74c3c') if IMPRESSAO_COLORIDA else colors.HexColor('#555555')
        msgs.append("ALERTA: ACIMA DO MÁXIMO")
    else:
        if minimo is not None and atual < minimo:
            cor = colors.HexColor('#e67e22') if IMPRESSAO_COLORIDA else colors.HexColor('#777777')
            msgs.append("ALERTA: ESTOQUE BAIXO")
        else:
            cor = colors.HexColor('#2ecc71') if IMPRESSAO_COLORIDA else colors.HexColor('#999999')
        # Usa ideal como teto de envio quando disponível; caso contrário usa máximo
        teto_envio = ideal if ideal is not None else maximo
        enviar = teto_envio - atual
        if enviar > 0:
            msgs.append(f"ENVIAR NO MÁXIMO {formatar_numero(enviar)} PEÇAS")
    return cor, msgs


def status_modelagem_ml(df_extrato, filial, ml_trad, ml_total):
    if ml_total == 0:
        return []
    
    pct_trad_loja = (ml_trad / ml_total * 100) if ml_total > 0 else 0
    
    cols_tam = [c for c in df_extrato.columns if c.isdigit()]
    df_ml = df_extrato[
        (df_extrato['Subgrupo'].str.contains('MANGA LONGA', na=False)) &
        (df_extrato['Tipo Produto'].isin(['ADULTO', 'ADULTO EXTRA']))
    ]
    
    total_ml_rede = df_ml[cols_tam].sum().sum()
    if total_ml_rede == 0:
        return []
    
    df_trad_rede = df_ml[
        (df_ml['Tipo Produto'] == 'ADULTO') &
        (df_ml['Subcategoria  Produto'] == 'TRADICIONAL')
    ]
    total_trad_rede = df_trad_rede[cols_tam].sum().sum()
    
    pct_trad_rede = (total_trad_rede / total_ml_rede * 100) if total_ml_rede > 0 else 0
    
    limite_min = pct_trad_rede / FATOR_ML_TRAD_MIN
    limite_max = pct_trad_rede * FATOR_ML_TRAD_MAX
    
    if pct_trad_loja > limite_max:
        return ["ALERTA: ESTOQUE DE TRADICIONAL EM EXCESSO"]
    elif pct_trad_loja < limite_min:
        return ["ALERTA: ESTOQUE TRADICIONAL ABAIXO DO NORMAL"]
    return []


def status_modelagem_mc(df_extrato, filial, mc_fit, mc_total):
    if mc_total == 0:
        return []
    
    pct_fit_loja = (mc_fit / mc_total * 100) if mc_total > 0 else 0
    
    cols_tam = [c for c in df_extrato.columns if c.isdigit()]
    df_mc = df_extrato[
        (df_extrato['Subgrupo'].str.contains('MANGA CURTA', na=False)) &
        (df_extrato['Tipo Produto'].isin(['ADULTO', 'ADULTO EXTRA']))
    ]
    
    total_mc_rede = df_mc[cols_tam].sum().sum()
    if total_mc_rede == 0:
        return []
    
    df_fit_rede = df_mc[
        (df_mc['Tipo Produto'] == 'ADULTO') &
        (df_mc['Subcategoria  Produto'] == 'SLIM')
    ]
    total_fit_rede = df_fit_rede[cols_tam].sum().sum()
    
    pct_fit_rede = (total_fit_rede / total_mc_rede * 100) if total_mc_rede > 0 else 0
    
    limite_min = pct_fit_rede / FATOR_MC_FIT_MIN
    limite_max = pct_fit_rede * FATOR_MC_FIT_MAX
    
    if pct_fit_loja > limite_max:
        return ["ALERTA: ESTOQUE DE FIT EM EXCESSO"]
    elif pct_fit_loja < limite_min:
        return ["ALERTA: ESTOQUE FIT ABAIXO DO NORMAL"]
    return []


HEADER_COLOR  = colors.HexColor('#34495e') if IMPRESSAO_COLORIDA else colors.black
HEADER2_COLOR = colors.HexColor('#5d6d7e') if IMPRESSAO_COLORIDA else colors.HexColor('#333333')
MODELS_HEADER = colors.HexColor('#2c3e50') if IMPRESSAO_COLORIDA else colors.HexColor('#222222')
TOTAL_BG      = colors.HexColor('#d5dbdb') if IMPRESSAO_COLORIDA else colors.HexColor('#cccccc')


def _estilo_estoque(cor_dados):
    return TableStyle([
        ('BACKGROUND',    (0,0),(-1,0), HEADER_COLOR),
        ('TEXTCOLOR',     (0,0),(-1,0), colors.whitesmoke),
        ('FONTNAME',      (0,0),(-1,0), 'Helvetica-Bold'),
        ('FONTSIZE',      (0,0),(-1,0), 11),
        ('BOTTOMPADDING', (0,0),(-1,0), 10),
        ('ALIGN',         (0,0),(-1,-1),'CENTER'),
        ('BACKGROUND',    (0,1),(-1,1), cor_dados),
        ('TEXTCOLOR',     (0,1),(-1,1), colors.whitesmoke),
        ('FONTNAME',      (0,1),(-1,1), 'Helvetica-Bold'),
        ('FONTSIZE',      (0,1),(-1,1), 11),
        ('GRID',          (0,0),(-1,-1), 1, colors.black),
    ])


def _estilo_modelagem():
    return TableStyle([
        ('BACKGROUND',    (0,0),(-1,0), HEADER2_COLOR),
        ('TEXTCOLOR',     (0,0),(-1,0), colors.whitesmoke),
        ('FONTNAME',      (0,0),(-1,0), 'Helvetica-Bold'),
        ('FONTSIZE',      (0,0),(-1,0), 10),
        ('BOTTOMPADDING', (0,0),(-1,0), 8),
        ('ALIGN',         (0,0),(-1,-1),'CENTER'),
        ('FONTNAME',      (0,1),(-1,-1),'Helvetica'),
        ('FONTSIZE',      (0,1),(-1,-1), 10),
        ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white, colors.white]),
        ('GRID',          (0,0),(-1,-1), 0.5, colors.HexColor('#aaaaaa')),
    ])


def _tabela_estoque(atual, maximo, minimo, ideal=None):
    cor, _ = status_estoque(atual, maximo, minimo)
    if ideal is not None:
        diferenca = atual - ideal
        dados = [
            ['Estoque Ideal', 'Estoque Atual', 'Diferença'],
            [formatar_numero(ideal), formatar_numero(atual),
             f"+{formatar_numero(diferenca)}" if diferenca >= 0 else formatar_numero(diferenca)],
        ]
        t = Table(dados, colWidths=[5.17*cm, 5.17*cm, 5.17*cm])
    else:
        min_str = formatar_numero(minimo) if minimo is not None else '—'
        diferenca = atual - maximo
        dados = [
            ['Estoque Máximo', 'Estoque Mínimo', 'Estoque Atual', 'Diferença'],
            [formatar_numero(maximo), min_str, formatar_numero(atual),
             f"+{formatar_numero(diferenca)}" if diferenca >= 0 else formatar_numero(diferenca)],
        ]
        t = Table(dados, colWidths=[4.25*cm, 3.75*cm, 3.75*cm, 3.75*cm])
    t.setStyle(_estilo_estoque(cor))
    return t


def _tabela_modelagem(fit, tradicional, tradicional_extra):
    dados = [
        ['Fit', 'Tradicional', 'Tradicional Extra'],
        [formatar_numero(fit), formatar_numero(tradicional), formatar_numero(tradicional_extra)]
    ]
    t = Table(dados, colWidths=[5.17*cm, 5.17*cm, 5.17*cm])
    t.setStyle(_estilo_modelagem())
    return t


def extrair_tamanhos_destacar(anotacao):
    """Extrai tamanhos que devem ser destacados das anotações.
    Suporta formato agrupado: 'POUCO TAM 02, 04 E 05' ou 'MUITO TAM 03'.
    """
    if not anotacao:
        return []
    # Captura todos os números após "POUCO TAM" ou "MUITO TAM" até fim do segmento
    tamanhos = re.findall(r'(?:POUCO|MUITO) TAM ([\d, E]+?)(?:\||$)', anotacao)
    resultado = []
    for grupo in tamanhos:
        resultado.extend(re.findall(r'\d+', grupo))
    return resultado


def _tabela_modelos(dados, label):
    linhas = dados['linhas']
    tam_cols = dados['tam_cols']
    total_tam = dados.get('total_tam', {})
    outros = dados.get('outros', None)
    
    if not linhas:
        return None
    
    PAGE_W = 17 * cm
    tem_descricao = any(l.get('descricao', '') for l in linhas)
    
    # Criar estilo para anotações com quebra de linha automática
    styles = getSampleStyleSheet()
    anotacao_style = ParagraphStyle(
        'Anotacao',
        parent=styles['Normal'],
        fontSize=7,
        fontName='Helvetica-Bold',
        leading=8,
        alignment=TA_LEFT,
        wordWrap='CJK'
    )
    
    # AJUSTE FINAL: Desc com pelo menos 2cm, tamanhos compactos, anotações máximas
    if tem_descricao:
        W_DESC = 3.2*cm; W_COD = 1.4*cm; W_COR = 0.5*cm; W_TOT = 0.8*cm; W_ANOT = 5.6*cm
        w_tam = max((PAGE_W - W_DESC - W_COD - W_COR - W_TOT - W_ANOT) / max(len(tam_cols), 1), 0.40*cm)
        col_widths = [W_DESC, W_COD, W_COR] + [w_tam]*len(tam_cols) + [W_TOT, W_ANOT]
        cabecalho = ['Desc', 'Cód', 'Cor'] + tam_cols + ['Tot', 'Anotações']
    else:
        W_COD = 1.4*cm; W_COR = 0.5*cm; W_TOT = 0.8*cm; W_ANOT = 7.8*cm
        w_tam = max((PAGE_W - W_COD - W_COR - W_TOT - W_ANOT) / max(len(tam_cols), 1), 0.45*cm)
        col_widths = [W_COD, W_COR] + [w_tam]*len(tam_cols) + [W_TOT, W_ANOT]
        cabecalho = ['Cód', 'Cor'] + tam_cols + ['Tot', 'Anotações']
    
    tabela_dados = [cabecalho]
    
    for l in linhas:
        tam_vals = [str(l['tamanhos'].get(t, 0)) if l['tamanhos'].get(t, 0) > 0 else '·'
                    for t in tam_cols]
        # Usar Paragraph para anotações para garantir quebra de linha
        anotacao_text = l['anotacao'] if l['anotacao'] else ""
        anotacao_para = Paragraph(anotacao_text, anotacao_style) if anotacao_text else ""
        
        if tem_descricao:
            row_data = [l.get('descricao', ''), l['produto'], l['cor']] + tam_vals + [str(l['total']), anotacao_para]
        else:
            row_data = [l['produto'], l['cor']] + tam_vals + [str(l['total']), anotacao_para]
        tabela_dados.append(row_data)
    
    idx_outros = None
    if outros and outros['total'] > 0:
        idx_outros = len(tabela_dados)
        tam_vals_outros = [str(outros['tamanhos'].get(t, 0)) if outros['tamanhos'].get(t, 0) > 0 else '·'
                           for t in tam_cols]
        if tem_descricao:
            tabela_dados.append([outros['label'], '—', '—'] + tam_vals_outros + [str(outros['total']), ''])
        else:
            tabela_dados.append([outros['label'], '—'] + tam_vals_outros + [str(outros['total']), ''])
    
    idx_total = len(tabela_dados)
    tam_vals_total = [str(total_tam.get(t, 0)) if total_tam.get(t, 0) > 0 else '·'
                      for t in tam_cols]
    total_geral = sum(total_tam.values())
    
    # CORREÇÃO 4: Incluir alertas do total da tabela
    alertas_total = dados.get('alertas_total', [])
    anotacao_total_text = " | ".join(alertas_total) if alertas_total else ""
    anotacao_total_para = Paragraph(anotacao_total_text, anotacao_style) if anotacao_total_text else ""
    
    if tem_descricao:
        tabela_dados.append([f'Total {label}', '—', '—'] + tam_vals_total + [str(total_geral), anotacao_total_para])
    else:
        tabela_dados.append([f'Total {label}', '—'] + tam_vals_total + [str(total_geral), anotacao_total_para])
    
    # Estilos básicos
    cmds = [
        ('BACKGROUND',    (0,0),(-1,0),  MODELS_HEADER),
        ('TEXTCOLOR',     (0,0),(-1,0),  colors.whitesmoke),
        ('FONTNAME',      (0,0),(-1,0),  'Helvetica-Bold'),
        ('FONTSIZE',      (0,0),(-1,0),  7),
        ('BOTTOMPADDING', (0,0),(-1,0),  6),
        ('TOPPADDING',    (0,0),(-1,0),  6),
        ('ALIGN',         (0,0),(-1,-1), 'CENTER'),
        ('ALIGN',         (0,1),(0,-1),  'LEFT'),
        ('ALIGN',         (-1,0),(-1,-1),'LEFT'),
        ('FONTNAME',      (0,1),(-1,-1), 'Helvetica'),
        ('FONTSIZE',      (0,1),(-1,-1), 7),
        ('FONTSIZE',      (-1,1),(-1,-1), 7),
        ('FONTNAME',      (-2,1),(-2,-1), 'Helvetica-Bold'),
        ('BACKGROUND',    (-2,0),(-2,0),  colors.HexColor('#1a252f') if IMPRESSAO_COLORIDA else colors.black),
        ('BACKGROUND',    (-2,1),(-2,idx_total-1), TOTAL_BG),
        ('GRID',          (0,0),(-1,-1),  0.4, colors.HexColor('#999999') if IMPRESSAO_COLORIDA else colors.black),
        ('VALIGN',        (0,0),(-1,-1), 'TOP'),
    ]
    
    # Background branco para linhas de produtos
    for i in range(1, idx_total):
        if i != idx_outros:
            cmds += [('BACKGROUND', (0,i),(-1,i), colors.white)]
    
    # Destacar células de tamanho mencionadas nas anotações
    col_offset = 3 if tem_descricao else 2
    
    for idx, l in enumerate(linhas):
        row_idx = idx + 1
        anotacao = l.get('anotacao', '')
        tamanhos_destacar = extrair_tamanhos_destacar(anotacao)
        
        for tam in tamanhos_destacar:
            if tam in tam_cols:
                col_idx = tam_cols.index(tam) + col_offset
                # Aplicar fundo cinza claro
                cmds.append(('BACKGROUND', (col_idx, row_idx), (col_idx, row_idx), 
                           colors.HexColor('#e0e0e0') if IMPRESSAO_COLORIDA else colors.HexColor('#d5d5d5')))
    
    if idx_outros:
        cmds += [
            ('BACKGROUND', (0,idx_outros),(-1,idx_outros), colors.white),
            ('FONTNAME',   (0,idx_outros),(-1,idx_outros), 'Helvetica-BoldOblique'),
        ]
    
    cmds += [
        ('BACKGROUND',  (0,idx_total),(-1,idx_total), colors.HexColor('#f0f0f0') if IMPRESSAO_COLORIDA else colors.HexColor('#dddddd')),
        ('FONTNAME',    (0,idx_total),(-1,idx_total), 'Helvetica-Bold'),
        ('TEXTCOLOR',   (0,idx_total),(-1,idx_total), colors.HexColor('#2c3e50') if IMPRESSAO_COLORIDA else colors.black),
    ]
    
    t = Table(tabela_dados, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle(cmds))
    return t


def _msgs_para_elementos(msgs, alerta_s, info_s):
    elems = []
    for msg in msgs:
        style = alerta_s if msg.startswith("ALERTA") else info_s
        prefix = "!  " if msg.startswith("ALERTA") else "→  "
        elems.append(Paragraph(f"{prefix}{msg}", style))
    return elems


def gerar_relatorio_pdf(filial, df_extrato, df_vendas, df_ideal_abs, df_compras, est, maximo, minimo, fator_3b=1.0, nome_arquivo=None):
    if nome_arquivo is None:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_arquivo = f"Relatorio_Estoque_{filial.replace(' ', '_')}_{ts}.pdf"
    
    doc = SimpleDocTemplate(nome_arquivo, pagesize=A4,
                            leftMargin=2*cm, rightMargin=2*cm,
                            topMargin=2*cm, bottomMargin=2*cm)
    e = []
    styles = getSampleStyleSheet()
    
    titulo_s = ParagraphStyle('Titulo', parent=styles['Heading1'],
        fontSize=18, textColor=colors.HexColor('#1f4788') if IMPRESSAO_COLORIDA else colors.black,
        spaceAfter=10, alignment=TA_CENTER, fontName='Helvetica-Bold')
    filial_s = ParagraphStyle('Filial', parent=styles['Heading1'],
        fontSize=14, textColor=colors.HexColor('#2c5aa0') if IMPRESSAO_COLORIDA else colors.black,
        spaceAfter=4, alignment=TA_CENTER, fontName='Helvetica-Bold')
    sub_s = ParagraphStyle('Sub', parent=styles['Heading2'],
        fontSize=13, textColor=colors.HexColor('#2c5aa0') if IMPRESSAO_COLORIDA else colors.black,
        spaceAfter=6, spaceBefore=4, alignment=TA_LEFT, fontName='Helvetica-Bold')
    subsub_s = ParagraphStyle('SubSub', parent=styles['Normal'],
        fontSize=11, textColor=colors.HexColor('#5d6d7e') if IMPRESSAO_COLORIDA else colors.black,
        spaceAfter=4, fontName='Helvetica-Bold')
    normal_s = ParagraphStyle('Norm', parent=styles['Normal'],
        fontSize=10, spaceAfter=8)
    alerta_s = ParagraphStyle('Alerta', parent=styles['Normal'],
        fontSize=11, textColor=colors.HexColor('#c0392b') if IMPRESSAO_COLORIDA else colors.black,
        spaceAfter=4, fontName='Helvetica-Bold')
    info_s = ParagraphStyle('Info', parent=styles['Normal'],
        fontSize=11, textColor=colors.HexColor('#1a5276') if IMPRESSAO_COLORIDA else colors.black,
        spaceAfter=4, fontName='Helvetica-Bold')
    modelos_s = ParagraphStyle('Modelos', parent=styles['Normal'],
        fontSize=10, textColor=colors.HexColor('#2c3e50') if IMPRESSAO_COLORIDA else colors.black,
        spaceAfter=4, spaceBefore=6, fontName='Helvetica-Bold')
    
    compras = obter_compras_a_receber(df_compras, filial)
    ideal_vals = obter_ideal(df_ideal_abs, filial)
    
    e.append(Paragraph("RELATÓRIO DE ESTOQUE DE CAMISAS", titulo_s))
    e.append(Paragraph(filial, filial_s))
    e.append(Spacer(1, 0.2*cm))
    e.append(Paragraph(f"<i>Gerado em: {datetime.now().strftime('%d/%m/%Y às %H:%M')}</i>", normal_s))
    e.append(Spacer(1, 0.6*cm))
    
    # 1. MANGA LONGA
    ml_atual = est['MANGA LONGA_ADULTO']
    ml_max = maximo['MANGA LONGA'];  ml_min = minimo['MANGA LONGA']
    _, msgs_ml = status_estoque(ml_atual, ml_max, ml_min, ideal=ideal_vals["MANGA LONGA"] if ideal_vals else None)
    msgs_mod_ml = status_modelagem_ml(df_extrato, filial, est['ML_TRADICIONAL'], ml_atual)
    
    msgs_ml.extend(compras['ML'])
    
    e.append(Paragraph("1. MANGA LONGA (Adulto + Adulto Extra)", sub_s))
    e.append(_tabela_estoque(ml_atual, ml_max, ml_min, ideal=ideal_vals["MANGA LONGA"] if ideal_vals else None))
    e.append(Spacer(1, 0.25*cm))
    for el in _msgs_para_elementos(msgs_ml, alerta_s, info_s): e.append(el)
    e.append(Spacer(1, 0.2*cm))
    e.append(Paragraph("Estoque por Modelagem:", subsub_s))
    e.append(_tabela_modelagem(est['ML_FIT'], est['ML_TRADICIONAL'], est['ML_TRADICIONAL_EXTRA']))
    for el in _msgs_para_elementos(msgs_mod_ml, alerta_s, info_s): e.append(el)
    
    dados_ml_fit = montar_tabela_subcat(df_extrato, df_vendas, filial,
                                         'MANGA LONGA', ['ADULTO', 'ADULTO EXTRA'], 'SLIM', TAM_ML_FIT, fator_3b=fator_3b)
    t_ml_fit = _tabela_modelos(dados_ml_fit, 'Fit')
    if t_ml_fit:
        e.append(Spacer(1, 0.3*cm))
        e.append(Paragraph("Estoque Modelos Fit:", modelos_s))
        e.append(t_ml_fit)
    
    dados_ml_trad = montar_tabela_subcat(df_extrato, df_vendas, filial,
                                          'MANGA LONGA', ['ADULTO', 'ADULTO EXTRA'], 'TRADICIONAL', TAM_ML_TRADICIONAL, fator_3b=fator_3b)
    t_ml_trad = _tabela_modelos(dados_ml_trad, 'Tradicional')
    if t_ml_trad:
        e.append(Spacer(1, 0.3*cm))
        e.append(Paragraph("Estoque Modelos Tradicional:", modelos_s))
        e.append(t_ml_trad)
    
    e.append(Spacer(1, 0.8*cm))
    
    # 2. MANGA CURTA
    mc_atual = est['MANGA CURTA_ADULTO']
    mc_max = maximo['MANGA CURTA'];  mc_min = minimo['MANGA CURTA']
    _, msgs_mc = status_estoque(mc_atual, mc_max, mc_min, ideal=ideal_vals["MANGA CURTA"] if ideal_vals else None)
    msgs_mod_mc = status_modelagem_mc(df_extrato, filial, est['MC_FIT'], mc_atual)
    
    msgs_mc.extend(compras['MC'])
    
    e.append(Paragraph("2. MANGA CURTA (Adulto + Adulto Extra)", sub_s))
    e.append(_tabela_estoque(mc_atual, mc_max, mc_min, ideal=ideal_vals["MANGA CURTA"] if ideal_vals else None))
    e.append(Spacer(1, 0.25*cm))
    for el in _msgs_para_elementos(msgs_mc, alerta_s, info_s): e.append(el)
    e.append(Spacer(1, 0.2*cm))
    e.append(Paragraph("Estoque por Modelagem:", subsub_s))
    e.append(_tabela_modelagem(est['MC_FIT'], est['MC_TRADICIONAL'], est['MC_TRADICIONAL_EXTRA']))
    for el in _msgs_para_elementos(msgs_mod_mc, alerta_s, info_s): e.append(el)
    
    dados_mc_fit = montar_tabela_subcat(df_extrato, df_vendas, filial,
                                         'MANGA CURTA', ['ADULTO', 'ADULTO EXTRA'], 'SLIM', TAM_MC_ADULTO, fator_3b=fator_3b)
    t_mc_fit = _tabela_modelos(dados_mc_fit, 'Fit')
    if t_mc_fit:
        e.append(Spacer(1, 0.3*cm))
        e.append(Paragraph("Estoque Modelos Fit:", modelos_s))
        e.append(t_mc_fit)
    
    dados_mc_trad = montar_tabela_subcat(df_extrato, df_vendas, filial,
                                          'MANGA CURTA', ['ADULTO', 'ADULTO EXTRA'], 'TRADICIONAL', TAM_MC_ADULTO, fator_3b=fator_3b)
    t_mc_trad = _tabela_modelos(dados_mc_trad, 'Tradicional')
    if t_mc_trad:
        e.append(Spacer(1, 0.3*cm))
        e.append(Paragraph("Estoque Modelos Tradicional:", modelos_s))
        e.append(t_mc_trad)
    
    e.append(Spacer(1, 0.8*cm))
    
    # 3. JUVENIL
    juv_atual = est['JUVENIL_TOTAL']
    juv_max = maximo['JUVENIL'];  juv_min = minimo['JUVENIL']
    _, msgs_juv = status_estoque(juv_atual, juv_max, juv_min, ideal=ideal_vals["JUVENIL"] if ideal_vals else None)
    
    msgs_juv.extend(compras['JUV'])
    
    e.append(Paragraph("3. JUVENIL", sub_s))
    e.append(_tabela_estoque(juv_atual, juv_max, juv_min, ideal=ideal_vals["JUVENIL"] if ideal_vals else None))
    e.append(Spacer(1, 0.25*cm))
    for el in _msgs_para_elementos(msgs_juv, alerta_s, info_s): e.append(el)
    
    dados_juv = montar_tabela_juvenil(df_extrato, df_vendas, filial, TAM_JUVENIL, fator_3b=fator_3b)
    t_juv = _tabela_modelos(dados_juv, 'Juvenil')
    if t_juv:
        e.append(Spacer(1, 0.3*cm))
        e.append(Paragraph("Estoque Modelos Juvenil:", modelos_s))
        e.append(t_juv)
    
    doc.build(e)
    return nome_arquivo


def resolver_filial(escolha, lojas, mapa_numero):
    """Resolve nome da filial a partir de número ou nome completo."""
    escolha = escolha.strip()
    if escolha.isdigit():
        chave = str(int(escolha))
        if chave in mapa_numero:
            return mapa_numero[chave]
        return None
    nome = escolha.upper().strip()
    return nome if nome in lojas else None


def main():
    print("=" * 70)
    print("GERADOR DE RELATÓRIO DE ESTOQUE - EXPEDIÇÃO")
    print("=" * 70)
    print()

    df_extrato, df_max, df_min, df_vendas, df_ideal_abs, df_compras, fator_3b = carregar_dados()
    lojas = listar_lojas_disponiveis(df_extrato, df_max)

    if not lojas:
        print("ERRO Nenhuma loja encontrada!")
        sys.exit(1)

    print(f"\nOK {len(lojas)} lojas disponíveis\n")

    mapa_numero = {}
    for loja in lojas:
        m = re.search(r'(\d+)', loja)
        if m:
            mapa_numero[str(int(m.group(1)))] = loja

    # ── Modo automático: loja passada como argumento (ex: python script.py 56) ──
    if len(sys.argv) > 1:
        escolha = sys.argv[1]
        filial = resolver_filial(escolha, lojas, mapa_numero)
        if filial is None:
            print(f"ERRO Loja '{escolha}' não encontrada ou inativa.")
            sys.exit(1)

    # ── Modo interativo: pergunta ao usuário ──────────────────────────────────
    else:
        print("LOJAS DISPONÍVEIS:")
        print("-" * 70)
        for loja in lojas:
            print(f"  {loja}")
        print("-" * 70)

        while True:
            try:
                print("\nDigite o número da loja (ex: 32) ou nome completo (ex: LOJA 32): ", end="")
                escolha = input().strip()
                filial = resolver_filial(escolha, lojas, mapa_numero)
                if filial:
                    break
                else:
                    if escolha.isdigit():
                        print(f"ERRO Loja {escolha} não encontrada ou está inativa.")
                    else:
                        print(f"ERRO Loja '{escolha}' não encontrada.")
            except KeyboardInterrupt:
                print("\n\nOperação cancelada.")
                sys.exit(0)
    
    print(f"\n{'='*70}")
    print(f"Gerando relatório para: {filial}")
    print(f"{'='*70}\n")
    
    est = calcular_estoque_loja(df_extrato, filial)
    resultado_limites = obter_limites(df_max, df_min, filial)
    
    if est is None or resultado_limites is None:
        print("ERRO Dados não encontrados.")
        sys.exit(1)
    
    maximo, minimo = resultado_limites
    
    try:
        nome_pdf = gerar_relatorio_pdf(filial, df_extrato, df_vendas, df_ideal_abs, df_compras, est, maximo, minimo, fator_3b=fator_3b)
        print(f"OK Relatório gerado: {nome_pdf}")
        print(f"\n{'='*70}")
        print("CONCLUÍDO!")
        print(f"{'='*70}")
    except Exception as ex:
        import traceback
        print(f"ERRO Erro ao gerar PDF: {ex}")
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
