import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
import re
import pandas as pd
import numpy as np
from pathlib import Path

from html.parser import HTMLParser

class _TableHTMLParser(HTMLParser):
    """Parser simples (stdlib) para extrair a PRIMEIRA tabela de um HTML.
    Usado como fallback quando pandas.read_html exige dependências opcionais (bs4/lxml)."""
    def __init__(self):
        super().__init__()
        self.in_table = False
        self.in_tr = False
        self.in_cell = False
        self._cell_buf = []
        self._row = []
        self._rows = []
        self.tables = []
        self._table_depth = 0

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        if tag == "table":
            self._table_depth += 1
            if not self.in_table:
                self.in_table = True
                self._rows = []
        if self.in_table and tag == "tr":
            self.in_tr = True
            self._row = []
        if self.in_table and self.in_tr and tag in ("td", "th"):
            self.in_cell = True
            self._cell_buf = []

    def handle_endtag(self, tag):
        tag = tag.lower()
        if self.in_table and self.in_tr and tag in ("td", "th") and self.in_cell:
            txt = "".join(self._cell_buf).strip()
            txt = re.sub(r"\s+", " ", txt)
            self._row.append(txt)
            self.in_cell = False
            self._cell_buf = []
        if self.in_table and tag == "tr" and self.in_tr:
            if len(self._row) > 0 and any(c != "" for c in self._row):
                self._rows.append(self._row)
            self.in_tr = False
            self._row = []
        if tag == "table" and self.in_table:
            self._table_depth -= 1
            if self._table_depth <= 0:
                self.tables.append(self._rows)
                self.in_table = False
                self._table_depth = 0
                self._rows = []

    def handle_data(self, data):
        if self.in_cell:
            self._cell_buf.append(data)

def _read_first_html_table_stdlib(file_path: Path, encoding: str):
    raw = file_path.read_bytes()
    try:
        text = raw.decode(encoding, errors="replace")
    except LookupError:
        text = raw.decode("latin1", errors="replace")

    p = _TableHTMLParser()
    p.feed(text)
    if not p.tables:
        raise ValueError(f"Nenhuma tabela HTML encontrada em {file_path}")

    rows = p.tables[0]
    maxlen = max((len(r) for r in rows), default=0)
    rect = [r + [""] * (maxlen - len(r)) for r in rows]
    return pd.DataFrame(rect)


# Arquivos de entrada (arquivos HTML salvos como .XLS)
ARQ_MANGA_CURTA = "MANGA_CURTA.XLS"
ARQ_MANGA_LONGA = "MANGA_LONGA.XLS"
ARQ_COMPRAS = "COMPRAS.XLS"

# Arquivo de saída
ARQ_SAIDA_COMBINADO = "ESTOQUE_AJUSTADO.xlsx"
ARQ_SAIDA_COMPRAS = "COMPRAS_AJUSTADO.xlsx"


def _auto_detect_header_row(df_raw: pd.DataFrame, required_terms):
    """
    Detecta a linha de cabeçalho em um DataFrame lido sem header,
    procurando uma linha que contenha (em qualquer ordem) todos os termos exigidos.
    """
    req = [t.strip().lower() for t in required_terms]

    for i, row in df_raw.iterrows():
        cells = []
        for cell in row.tolist():
            if isinstance(cell, str):
                cells.append(cell.strip().lower())
            else:
                cells.append("")
        joined = " | ".join(cells)
        if all(term in joined for term in req):
            return i
    return None


def ler_arquivo_html_xls(file_path, encoding='ISO-8859-1', header_hint=None):
    """
    Lê um arquivo exportado como HTML e salvo com extensão .XLS (muito comum em sistemas legados).
    Retorna um DataFrame com o cabeçalho na linha correta.

    Observação importante:
    - pd.read_html pode exigir dependências opcionais (beautifulsoup4/lxml). Para evitar isso,
      este leitor tenta pd.read_html e, se faltar dependência, cai para um parser da stdlib.
    """
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {file_path}")

    def _read_html_raw_noheader():
        # 1) tenta pandas.read_html
        try:
            tables = pd.read_html(file_path, encoding=encoding, header=None)
            if not tables:
                raise ValueError("Nenhuma tabela encontrada")
            return tables[0]
        except ImportError:
            # 2) fallback stdlib (não depende de bs4)
            return _read_first_html_table_stdlib(file_path, encoding=encoding)

    df_raw = _read_html_raw_noheader()

    # Se tiver um hint de header (ex: MANGA_* normalmente header=3), aplica direto
    if header_hint is not None:
        hdr = int(header_hint)
    else:
        # Auto-detect do cabeçalho
        hdr = _auto_detect_header_row(df_raw, required_terms=["filial", "produto", "grade"])
        if hdr is None:
            hdr = _auto_detect_header_row(df_raw, required_terms=["filial", "fornecedor", "produto", "qtde"])
        if hdr is None:
            raise ValueError(f"Não consegui detectar a linha de cabeçalho em {file_path}")

    header = df_raw.iloc[hdr].tolist()
    df = df_raw.iloc[hdr + 1:].copy()
    df.columns = header
    df = df.reset_index(drop=True)

    # Remove colunas Unnamed
    mask = ~pd.Index(df.columns).astype(str).str.match(r"^Unnamed", na=False)
    df = df.loc[:, mask]

    # Padronização pontual de Produto (quando existir)
    if "Produto" in df.columns:
        df["Produto"] = df["Produto"].astype(str).str.replace(r"\.0$", "", regex=True)

    print(f"OK Arquivo {file_path.name} lido com sucesso: {df.shape[0]} linhas, {df.shape[1]} colunas")
    return df
def normalizar_filial(valor):
    """
    Normaliza o nome da filial:
    - "LOJA 06 - SHOP CPO LIMPO" -> "LOJA 06"
    - "LOJA 53 - BOM RETIRO" -> "LOJA 53"
    - "SANTA RITA" -> "SANTA RITA"
    - "E-COMMERCE" -> "E-COMMERCE"
    """
    if pd.isna(valor):
        return valor

    s = str(valor).strip().upper()

    # Casos especiais
    if s == "SANTA RITA":
        return "SANTA RITA"
    if s == "E-COMMERCE":
        return "E-COMMERCE"

    # Extrai número da loja
    m = re.match(r"^(LOJA)\s*(\d+)", s)
    if m:
        num = int(m.group(2))          # remove zeros à esquerda
        return f"LOJA {num:02d}"       # força 2 dígitos

    return s


def parse_grade_info(grade_value: str):
    """
    Espera algo como '01 - 07', '2 - 16', '02 - 06', '34 - 48', etc.
    Retorna (ini, fim, step) como ints.
    """
    if pd.isna(grade_value):
        return (np.nan, np.nan, np.nan)

    s = str(grade_value).strip()
    m = re.search(r"(\d+)\s*-\s*(\d+)", s)
    if not m:
        return (np.nan, np.nan, np.nan)

    ini = int(m.group(1))
    fim = int(m.group(2))

    # Determinar step
    if ini <= 2 and fim >= 16:
        step = 2  # Grade juvenil: 2-16 = 02, 04, 06, 08, 10, 12, 14, 16
    else:
        step = 1  # Grade normal sequencial

    return (ini, fim, step)


def reposicionar_es_para_tamanhos(
    df: pd.DataFrame,
    col_grade="Grade",
    es_prefix="Es",
    tam_min=1,
    tam_max=16
) -> pd.DataFrame:
    """
    Converte Es1..Es8 (posição relativa) em colunas de tamanho real,
    usando o início da grade e o step correto.
    """
    es_cols = [c for c in df.columns if str(c).startswith(es_prefix)]
    if not es_cols:
        raise ValueError("Não encontrei colunas Es1..Es8 no dataframe.")

    def processar_valor_es(valor):
        """
        Processa um valor da coluna Es, detectando o formato.
        - Se contém vírgula: formato decimal brasileiro (1.499,00 → 1499)
        - Se não contém vírgula: formato inteiro (95900 → 959)
        """
        valor_str = str(valor).strip()

        if valor_str in ['', 'nan', 'None', '000', '0']:
            return 0

        if ',' in valor_str:
            valor_str = valor_str.replace('.', '')
            valor_str = valor_str.replace(',00', '')
            valor_str = valor_str.replace(',', '.')
            try:
                return float(valor_str)
            except Exception:
                return 0
        else:
            valor_str = valor_str.replace('.', '')
            try:
                return float(valor_str) / 100
            except Exception:
                return 0

    for c in es_cols:
        df[c] = df[c].apply(processar_valor_es)

    grade_info = df[col_grade].apply(parse_grade_info)
    df["_grade_ini"] = grade_info.apply(lambda x: x[0])
    df["_grade_fim"] = grade_info.apply(lambda x: x[1])
    df["_grade_step"] = grade_info.apply(lambda x: x[2])

    df2 = df.reset_index(drop=False).rename(columns={"index": "_row_id"})

    long = df2.melt(
        id_vars=["_row_id", "_grade_ini", "_grade_fim", "_grade_step"],
        value_vars=es_cols,
        var_name="_es_col",
        value_name="_qtde"
    )

    long["_es_pos"] = long["_es_col"].str.extract(r"(\d+)").astype(float).astype("Int64")
    long["_tam_real"] = long["_grade_ini"] + ((long["_es_pos"] - 1) * long["_grade_step"])

    long = long[
        long["_tam_real"].between(tam_min, tam_max, inclusive="both") &
        (long["_tam_real"] <= long["_grade_fim"]) &
        long["_tam_real"].notna()
    ].copy()

    pivot = long.pivot_table(
        index="_row_id",
        columns="_tam_real",
        values="_qtde",
        aggfunc="sum",
        fill_value=0
    ).reset_index()

    pivot = pivot.rename(columns={
        c: f"{int(c):02d}" for c in pivot.columns
        if c != "_row_id" and isinstance(c, (int, float))
    })

    out = df2.merge(pivot, on="_row_id", how="left")

    for t in range(tam_min, tam_max + 1):
        col = f"{t:02d}"
        if col not in out.columns:
            out[col] = 0
        out[col] = out[col].fillna(0).round(0).astype(int)

    out = out.drop(columns=["_row_id", "_grade_ini", "_grade_fim", "_grade_step"], errors="ignore")
    return out


def processar_estoque(arquivo_entrada):
    """
    Processa um arquivo de estoque (MANGA_CURTA ou MANGA_LONGA).
    """
    print(f"\n{'='*60}")
    print(f"Processando ESTOQUE: {arquivo_entrada}")
    print(f"{'='*60}")

    df = ler_arquivo_html_xls(arquivo_entrada, header_hint=3)

    print(f"Colunas encontradas: {df.columns.tolist()}")

    colunas_para_preencher = ["Filial", "Tipo Produto", "Subgrupo"]
    for col in colunas_para_preencher:
        if col in df.columns:
            df[col] = df[col].replace(r"^\s*$", np.nan, regex=True)
            df[col] = df[col].ffill()
            print(f"OK Coluna '{col}' preenchida com forward fill")
        else:
            print(f"! Aviso: Coluna '{col}' não encontrada no arquivo")

    if "Filial" in df.columns:
        df["Filial"] = df["Filial"].apply(normalizar_filial)
        print("OK Nomes das filiais normalizados")

    if "Produto" in df.columns:
        df["Produto"] = df["Produto"].replace(r"^\s*$", np.nan, regex=True).ffill()
        print("OK Coluna 'Produto' preenchida")

    if "Grade" in df.columns:
        grade_norm = df["Grade"].astype(str).str.strip().str.upper()
        linhas_antes = len(df)
        df = df[grade_norm != "U"].copy()
        linhas_removidas = linhas_antes - len(df)
        if linhas_removidas > 0:
            print(f"OK Removidas {linhas_removidas} linhas com grade 'U' (tamanho único)")

    if "Tipo Produto" in df.columns:
        tipo_norm = df["Tipo Produto"].astype(str).str.strip().str.upper()
        linhas_antes = len(df)
        df = df[tipo_norm != "FEMININO"].copy()
        linhas_removidas = linhas_antes - len(df)
        if linhas_removidas > 0:
            print(f"OK Removidas {linhas_removidas} linhas de produtos FEMININOS")

    if "Filial" in df.columns:
        filiais_inativas = [
            "MATRIZ", "LOJA 01", "LOJA 02", "LOJA 08", "LOJA 10", "LOJA 11",
            "LOJA 12", "LOJA 15", "LOJA 18", "LOJA 19", "LOJA 20", "LOJA 22",
            "LOJA 24", "LOJA 25", "LOJA 27", "LOJA 30", "LOJA 35", "LOJA 43",
            "TOTAL GERAL"
        ]
        linhas_antes = len(df)
        df = df[~df["Filial"].isin(filiais_inativas)].copy()
        linhas_removidas = linhas_antes - len(df)
        if linhas_removidas > 0:
            print(f"OK Removidas {linhas_removidas} linhas de filiais INATIVAS e TOTAL GERAL")

    print("Convertendo colunas Es1..Es8 para tamanhos reais...")
    df_out = reposicionar_es_para_tamanhos(df, col_grade="Grade")

    es_cols_final = [c for c in df_out.columns if str(c).startswith("Es")]
    df_out = df_out.drop(columns=es_cols_final, errors="ignore")
    print("OK Colunas Es removidas, mantendo apenas tamanhos reais")

    size_cols = sorted([c for c in df_out.columns if re.match(r"^\d{2}$", str(c)) and int(c) <= 16])
    base_cols = [c for c in df_out.columns if c not in size_cols]
    df_out = df_out[base_cols + size_cols]

    print(f"OK Processamento ESTOQUE concluído: {len(df_out)} linhas, {len(df_out.columns)} colunas")
    return df_out


def processar_compras(arquivo_entrada):
    """
    Processa o arquivo COMPRAS.XLS e gera um DataFrame padronizado para uma aba separada no Excel.

    Estrutura esperada (padrão Linx/Expedição):
    - Filial A Entregar
    - Fornecedor
    - Produto
    - Descrição do  Produto
    - Qtde Entregar
    """
    print(f"\n{'='*60}")
    print(f"Processando COMPRAS: {arquivo_entrada}")
    print(f"{'='*60}")

    df = ler_arquivo_html_xls(arquivo_entrada, header_hint=None)
    print(f"Colunas encontradas: {df.columns.tolist()}")

    # Renomear colunas para padronizar
    ren = {}
    for c in df.columns:
        c_norm = str(c).strip().lower()
        if "filial" in c_norm:
            ren[c] = "Filial"
        elif "fornecedor" in c_norm:
            ren[c] = "Fornecedor"
        elif c_norm == "produto":
            ren[c] = "Produto"
        elif "descrição" in c_norm or "descricao" in c_norm:
            ren[c] = "Desc Produto"
        elif "qtde" in c_norm or "qtd" in c_norm:
            ren[c] = "Qtde Entregar"

    df = df.rename(columns=ren)

    # Normalizar filial
    if "Filial" in df.columns:
        df["Filial"] = df["Filial"].apply(normalizar_filial)

    # Produto como string sem .0
    if "Produto" in df.columns:
        df["Produto"] = df["Produto"].astype(str).str.replace(r"\.0$", "", regex=True)

    # Quantidade -> int (pode vir como '10,00' ou '1.499,00' etc.)
    # Importante: NÃO remover vírgula antes de interpretar, senão '10,00' vira '1000'.
    def _parse_qtde_ptbr(v):
        s = str(v).strip()
        if s in ("", "nan", "None"):
            return 0
        # remove espaços
        s = re.sub(r"\s+", "", s)
        # se tem vírgula, assume decimal pt-BR (milhares com ponto, decimais com vírgula)
        if "," in s:
            s2 = s.replace(".", "").replace(",", ".")
            try:
                return int(round(float(s2)))
            except Exception:
                return 0
        # sem vírgula: tenta lógica do estoque (alguns exports vêm em centavos)
        s2 = s.replace(".", "")
        # mantém só dígitos e sinal
        s2 = re.sub(r"[^0-9\-]", "", s2)
        if s2 in ("", "-", "0", "000"):
            return 0
        try:
            # se tiver cara de centavos (termina com 00 e tem pelo menos 3 dígitos), divide por 100
            if len(s2) >= 3 and s2.endswith("00"):
                return int(round(float(s2) / 100))
            return int(round(float(s2)))
        except Exception:
            return 0

    if "Qtde Entregar" in df.columns:
        df["Qtde Entregar"] = df["Qtde Entregar"].apply(_parse_qtde_ptbr).astype(int)

        # Remover linhas com qtde = 0
        linhas_antes = len(df)
        df = df[df["Qtde Entregar"] > 0].copy()
        print(f"OK Removidas {linhas_antes - len(df)} linhas com Qtde Entregar = 0")

    # Remover linhas totalmente vazias
    df = df.dropna(how="all")

    # Organizar colunas mais úteis primeiro
    col_order = [c for c in ["Filial", "Fornecedor", "Produto", "Desc Produto", "Qtde Entregar"] if c in df.columns]
    restantes = [c for c in df.columns if c not in col_order]
    df = df[col_order + restantes]

    print(f"OK Processamento COMPRAS concluído: {len(df)} linhas, {len(df.columns)} colunas")
    return df


def main():
    """
    Função principal:
    - Processa MANGA_CURTA e MANGA_LONGA e gera "ESTOQUE_AJUSTADO.xlsx" (aba "Estoque_Completo")
    - Processa COMPRAS e gera "COMPRAS_AJUSTADO.xlsx" (aba "Compras")
    """
    print("="*60)
    print("SISTEMA DE PADRONIZAÇÃO - ESTOQUE + COMPRAS (EXPEDIÇÃO)")
    print("="*60)

    # ---------------------------
    # 1) ESTOQUE (MANGAS)
    # ---------------------------
    dfs_estoque = []

    try:
        df_curta = processar_estoque(ARQ_MANGA_CURTA)
        df_curta["Origem"] = "MANGA_CURTA"
        dfs_estoque.append(df_curta)
    except Exception as e:
        print(f"\nERRO Erro ao processar {ARQ_MANGA_CURTA}: {e}")
        import traceback
        traceback.print_exc()

    try:
        df_longa = processar_estoque(ARQ_MANGA_LONGA)
        df_longa["Origem"] = "MANGA_LONGA"
        dfs_estoque.append(df_longa)
    except Exception as e:
        print(f"\nERRO Erro ao processar {ARQ_MANGA_LONGA}: {e}")
        import traceback
        traceback.print_exc()

    if dfs_estoque:
        out_estoque = Path(ARQ_SAIDA_COMBINADO)
        if out_estoque.exists():
            try:
                out_estoque.unlink()
            except Exception:
                pass

        df_estoque = pd.concat(dfs_estoque, ignore_index=True)
        with pd.ExcelWriter(ARQ_SAIDA_COMBINADO, engine="openpyxl") as writer:
            df_estoque.to_excel(writer, index=False, sheet_name="Estoque_Completo")

        print(f"OK Estoque gravado em '{ARQ_SAIDA_COMBINADO}' ({len(df_estoque)} linhas)")
    else:
        print("\n! Nenhum dataframe de estoque foi gerado. Verifique MANGA_CURTA/MANGA_LONGA.")

    # ---------------------------
    # 2) COMPRAS (ARQUIVO SEPARADO)
    # ---------------------------
    try:
        df_compras = processar_compras(ARQ_COMPRAS)

        out_compras = Path(ARQ_SAIDA_COMPRAS)
        if out_compras.exists():
            try:
                out_compras.unlink()
            except Exception:
                pass

        with pd.ExcelWriter(ARQ_SAIDA_COMPRAS, engine="openpyxl") as writer:
            df_compras.to_excel(writer, index=False, sheet_name="Compras")

        print(f"OK Compras gravado em '{ARQ_SAIDA_COMPRAS}' ({len(df_compras)} linhas)")
    except Exception as e:
        print(f"\nERRO Erro ao processar {ARQ_COMPRAS}: {e}")
        import traceback
        traceback.print_exc()

    print("\nPROCESSAMENTO CONCLUÍDO!")


if __name__ == "__main__":
    main()