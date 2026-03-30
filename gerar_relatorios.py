"""
ORQUESTRADOR DE RELATÓRIOS DE ESTOQUE
======================================
1. Pergunta qimport sys, io
if hasattr(sys.stdout, "buffer"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
uais lojas gerar
2. Roda padronizar_planilhas.py  (gera ESTOQUE_AJUSTADO e COMPRAS_AJUSTADO)
3. Roda gerar_relatorio_estoque.py para cada loja selecionada
"""

import sys
import re
import subprocess
from pathlib import Path
from datetime import datetime

SCRIPT_PADRONIZAR = "padronizar_planilhas.py"
SCRIPT_RELATORIO  = "gerar_relatorio_estoque.py"

# ─────────────────────────────────────────────────────────────────────────────
# Helpers de terminal
# ─────────────────────────────────────────────────────────────────────────────

def linha(char="=", n=70):
    print(char * n)

def titulo(texto):
    linha()
    print(texto)
    linha()

def rodar_script(script, args=None):
    """Executa um script Python com argumentos opcionais e retorna (ok, saída)."""
    cmd = [sys.executable, script] + (args or [])
    result = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        encoding='utf-8',
        errors='replace',
    )
    saida = (result.stdout or '') + (result.stderr or '')
    return result.returncode == 0, saida


# ─────────────────────────────────────────────────────────────────────────────
# Etapa 1 — descobrir lojas disponíveis
# ─────────────────────────────────────────────────────────────────────────────

def lojas_disponiveis():
    """
    Lê ESTOQUE_AJUSTADO.xlsx (se já existir) para listar lojas,
    ou retorna lista vazia (o padronizar ainda não rodou).
    """
    arq = Path("ESTOQUE_AJUSTADO.xlsx")
    if not arq.exists():
        return []
    try:
        import pandas as pd
        df = pd.read_excel(arq)
        lojas = sorted(df["Filial"].dropna().unique())
        # Exclui Santa Rita e E-Commerce (não geram relatório)
        lojas = [l for l in lojas if re.search(r"LOJA\s+\d+", str(l))]
        return lojas
    except Exception:
        return []


def extrair_numeros(texto):
    """'14, 33, 56' ou '14 33 56' ou '14-33' → [14, 33, 56]."""
    return [int(n) for n in re.findall(r"\d+", texto)]


def selecionar_lojas(lojas_disponiveis):
    """Interação com o analista para escolher as lojas."""
    print()
    if lojas_disponiveis:
        print("Lojas disponíveis no último ESTOQUE_AJUSTADO:")
        linha("-")
        grupos = [lojas_disponiveis[i:i+10] for i in range(0, len(lojas_disponiveis), 10)]
        for g in grupos:
            print("  " + "  ".join(g))
        linha("-")

    print()
    print("Digite os números das lojas desejadas separados por vírgula ou espaço.")
    print("Exemplos:  14 33 56   |   14, 33, 56   |   ALL (todas)")
    print()

    while True:
        try:
            resp = input("Lojas → ").strip()
        except KeyboardInterrupt:
            print("\nCancelado.")
            sys.exit(0)

        if not resp:
            print("  !  Digite ao menos uma loja.")
            continue

        if resp.upper() in ("ALL", "TODAS", "TUDO"):
            if lojas_disponiveis:
                return lojas_disponiveis
            else:
                print("  !  Não há lojas carregadas ainda (rode após padronizar). Tente novamente.")
                continue

        nums = extrair_numeros(resp)
        if not nums:
            print("  !  Não entendi. Use números separados por vírgula ou espaço.")
            continue

        lojas_fmt = [f"LOJA {n:02d}" for n in nums]

        if lojas_disponiveis:
            invalidas = [l for l in lojas_fmt if l not in lojas_disponiveis]
            if invalidas:
                print(f"  !  Lojas não encontradas: {', '.join(invalidas)}. Tente novamente.")
                continue

        return lojas_fmt


# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────

def main():
    titulo("GERADOR DE RELATÓRIOS DE ESTOQUE — EXPEDIÇÃO")
    print(f"Início: {datetime.now().strftime('%d/%m/%Y às %H:%M')}")

    # Verificar scripts necessários
    for script in [SCRIPT_PADRONIZAR, SCRIPT_RELATORIO]:
        if not Path(script).exists():
            print(f"\nERRO Script não encontrado: {script}")
            print("   Certifique-se de que todos os arquivos estão na mesma pasta.")
            sys.exit(1)

    # ── Etapa 1: selecionar lojas ─────────────────────────────────────────────
    titulo("ETAPA 1 — SELEÇÃO DE LOJAS")
    lojas_pre = lojas_disponiveis()
    lojas_selecionadas = selecionar_lojas(lojas_pre)

    print()
    print(f"OK {len(lojas_selecionadas)} loja(s) selecionada(s):")
    print("  " + ", ".join(lojas_selecionadas))

    # ── Etapa 2: padronizar planilhas ─────────────────────────────────────────
    print()
    titulo("ETAPA 2 — PADRONIZAÇÃO DAS PLANILHAS")
    print(f"Rodando {SCRIPT_PADRONIZAR}...")
    print()

    ok, saida = rodar_script(SCRIPT_PADRONIZAR, args=[])
    print(saida)

    if not ok:
        print("ERRO Falha na padronização. Verifique os arquivos de entrada e tente novamente.")
        sys.exit(1)

    print("OK Padronização concluída.")

    # ── Etapa 3: gerar PDFs ───────────────────────────────────────────────────
    print()
    titulo("ETAPA 3 — GERAÇÃO DOS RELATÓRIOS PDF")

    total   = len(lojas_selecionadas)
    gerados = []
    falhas  = []

    for i, loja in enumerate(lojas_selecionadas, 1):
        # Extrai número da loja para passar como entrada ao script
        m = re.search(r"\d+", loja)
        num_loja = str(int(m.group())) if m else loja

        print(f"[{i}/{total}] Gerando relatório para {loja}...", end=" ", flush=True)

        ok, saida = rodar_script(SCRIPT_RELATORIO, args=[num_loja])

        # Encontrar nome do arquivo gerado na saída
        match_pdf = re.search(r"Relatório gerado: (\S+\.pdf)", saida)
        if ok and match_pdf:
            nome_pdf = match_pdf.group(1)
            gerados.append((loja, nome_pdf))
            print(f"OK  {nome_pdf}")
        else:
            falhas.append(loja)
            print("ERRO")
            # Mostrar erro resumido
            for linha_err in saida.splitlines():
                if "ERRO" in linha_err or "Erro" in linha_err or "Error" in linha_err:
                    print(f"   {linha_err.strip()}")

    # ── Resumo final ──────────────────────────────────────────────────────────
    print()
    titulo("RESUMO")
    print(f"Término: {datetime.now().strftime('%d/%m/%Y às %H:%M')}")
    print()
    print(f"OK Gerados com sucesso: {len(gerados)}/{total}")
    for loja, pdf in gerados:
        print(f"   {loja} → {pdf}")

    if falhas:
        print()
        print(f"ERRO Falhas: {len(falhas)}/{total}")
        for loja in falhas:
            print(f"   {loja}")

    print()
    linha()


if __name__ == "__main__":
    main()
