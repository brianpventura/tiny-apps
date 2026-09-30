"""
Relatório de Produtos — API V3 Tiny
Gera planilha Excel com SKU, Descrição, NCM e Código EAN de todos os produtos.

Estrutura em duas fases:
  Fase 1 — listagem paginada (traz SKU, descrição e EAN/gtin já na lista)
  Fase 2 — detalhe paralelizado para obter NCM (não vem na listagem)
"""

import sys
import os
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path

# adiciona a pasta src/ ao sys.path (sobe ate achar a pasta que contem 'core')
for _p in Path(__file__).resolve().parents:
    if (_p / 'core').is_dir():
        sys.path.insert(0, str(_p))
        break

import pandas as pd
from tqdm import tqdm
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from core.auth_tiny import resolver_tenant

from core.tiny_client import TinyClient


# ------------------------------------------------------------------ #
#  Helpers de extração                                                 #
# ------------------------------------------------------------------ #

def _extrair_sku(produto: dict) -> str:
    return str(produto.get("sku") or produto.get("codigo") or "").strip()


def _extrair_descricao(produto: dict) -> str:
    return str(produto.get("descricao") or produto.get("nome") or "").strip()


def _extrair_ean(produto: dict) -> str:
    """Na API v3 do Tiny, o EAN/GTIN aparece como campo 'gtin'."""
    return str(produto.get("gtin") or "").strip()


def _extrair_ncm(produto: dict) -> str:
    """NCM só aparece no detalhe (não na listagem paginada)."""
    return str(produto.get("ncm") or "").strip()


# ------------------------------------------------------------------ #
#  Coleta de dados                                                     #
# ------------------------------------------------------------------ #

def coletar_produtos(client: TinyClient) -> list[dict]:
    """Retorna lista de dicts {SKU, Descrição, NCM, Código EAN} para todos os produtos."""

    # ── Fase 1: listagem paginada ────────────────────────────────────
    pbar_lista: tqdm | None = None

    def _set_total(total):
        nonlocal pbar_lista
        pbar_lista = tqdm(
            total=total,
            desc="Fase 1/2 — listando    ",
            unit="prod",
            ncols=80,
            colour="cyan",
        )

    # Coletamos primeiro o que já vem na listagem (sku, descrição, ean).
    # Guardamos também o id para buscar o NCM no detalhe.
    bruto: list[dict] = []
    for prod in client.produtos_todos(on_total=_set_total):
        if pbar_lista is not None:
            pbar_lista.update(1)
        sku = _extrair_sku(prod)
        descricao = _extrair_descricao(prod)
        if not sku and not descricao:
            continue
        bruto.append({
            "id": prod.get("id"),
            "sku": sku,
            "descricao": descricao,
            "ean": _extrair_ean(prod),
        })

    if pbar_lista is not None:
        pbar_lista.close()

    print(f"\n{len(bruto)} produtos coletados — buscando NCM nos detalhes...\n")

    # ── Fase 2: detalhes em paralelo (apenas para o NCM) ─────────────
    _MAX_WORKERS = 4  # workers simultâneos — rate limiter controla a taxa global

    ncm_por_id: dict[int, str] = {}
    lock = threading.Lock()

    def _buscar_ncm(produto_id):
        try:
            detalhe = client.obter_produto(produto_id)
            if "data" in detalhe:
                detalhe = detalhe["data"]
            ncm = _extrair_ncm(detalhe)
            with lock:
                ncm_por_id[produto_id] = ncm
        except Exception as e:
            tqdm.write(f"  [AVISO] Produto {produto_id}: {e}")
            with lock:
                ncm_por_id[produto_id] = ""

    with tqdm(
        total=len(bruto),
        desc="Fase 2/2 — detalhes    ",
        unit="prod",
        ncols=80,
        colour="green",
    ) as pbar:
        with ThreadPoolExecutor(max_workers=_MAX_WORKERS) as executor:
            futures = [executor.submit(_buscar_ncm, p["id"]) for p in bruto if p["id"]]
            for future in as_completed(futures):
                try:
                    future.result()
                except Exception as e:
                    tqdm.write(f"  [ERRO INESPERADO] {e}")
                pbar.update(1)

    # Junta os dados finais
    resultado = []
    for p in bruto:
        resultado.append({
            "SKU": p["sku"],
            "Descrição": p["descricao"],
            "NCM": ncm_por_id.get(p["id"], ""),
            "Código EAN": p["ean"],
        })

    print()
    return resultado


# ------------------------------------------------------------------ #
#  Geração da planilha                                                 #
# ------------------------------------------------------------------ #

def gerar_excel(produtos: list[dict], caminho: str) -> None:
    if not produtos:
        print("Nenhum produto encontrado.")
        return

    df = pd.DataFrame(produtos, columns=["SKU", "Descrição", "NCM", "Código EAN"])
    df = df.sort_values("Descrição").reset_index(drop=True)

    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Produtos")
        ws = writer.sheets["Produtos"]
        _formatar_planilha(ws, df)

    print(f"\nPlanilha salva em: {caminho}")
    print(f"Total: {len(df)} produtos")


def _formatar_planilha(ws, df: pd.DataFrame) -> None:
    COR_CABECALHO = "1F3864"
    BORDA_FINA = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )

    # Cabeçalho
    for col_idx, _ in enumerate(df.columns, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
        cell.fill = PatternFill("solid", fgColor=COR_CABECALHO)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDA_FINA

    # Corpo
    for row_idx in range(2, len(df) + 2):
        for col_idx, col_name in enumerate(df.columns, start=1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = Font(name="Calibri", size=11)
            cell.border = BORDA_FINA

            if col_name == "Descrição":
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")

            # Mantém SKU/EAN/NCM como texto (preserva zeros à esquerda e formato)
            if col_name in ("SKU", "Código EAN", "NCM"):
                cell.number_format = "@"

    # Larguras das colunas
    ws.column_dimensions["A"].width = 14   # SKU
    ws.column_dimensions["B"].width = 60   # Descrição
    ws.column_dimensions["C"].width = 14   # NCM
    ws.column_dimensions["D"].width = 18   # Código EAN
    ws.row_dimensions[1].height = 22


# ------------------------------------------------------------------ #
#  Ponto de entrada                                                    #
# ------------------------------------------------------------------ #

def main():
    print("=== Relatório de Produtos — Tiny API V3 ===\n")

    tenant_id = resolver_tenant(sys.argv[1] if len(sys.argv) > 1 else None)

    client = TinyClient(tenant_id)
    produtos = coletar_produtos(client)

    nome_arquivo = f"produtos_{tenant_id}_{datetime.now().strftime('%d%m%Y_%H%M')}.xlsx"
    caminho_arquivo = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__)))), "reports", nome_arquivo)
    gerar_excel(produtos, caminho_arquivo)


if __name__ == "__main__":
    main()
