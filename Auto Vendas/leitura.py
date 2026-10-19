"""Leitura dos arquivos Excel de entrada."""
from pathlib import Path
from zipfile import BadZipFile

import pandas as pd
from openpyxl.utils.exceptions import InvalidFileException

ARQUIVO_CLIENTES = "Clientes.xlsx"
ARQUIVO_ESTOQUE = "Estoque.xlsx"
ARQUIVO_PEDIDOS = "Pedidos.xlsx"


class ErroDeLeitura(Exception):
    """Falha ao localizar ou ler um arquivo de entrada."""


def _ler_planilha(caminho: Path, colunas_id: list[str]) -> pd.DataFrame:
    """Lê uma planilha preservando os identificadores como texto."""
    if not caminho.exists():
        raise ErroDeLeitura(
            f"Arquivo não encontrado: '{caminho.name}' (esperado em '{caminho.parent}')."
        )

    try:
        df = pd.read_excel(
            caminho,
            engine="openpyxl",
            dtype={coluna: str for coluna in colunas_id},
        )
    except PermissionError as erro:
        raise ErroDeLeitura(
            f"Sem permissão para ler '{caminho.name}'. Verifique se ele está aberto em outro programa."
        ) from erro
    except (BadZipFile, InvalidFileException, ValueError) as erro:
        raise ErroDeLeitura(
            f"'{caminho.name}' não é um arquivo .xlsx válido ou está corrompido ({erro})."
        ) from erro
    except OSError as erro:
        raise ErroDeLeitura(f"Não foi possível abrir '{caminho.name}': {erro}") from erro

    # Normalização inofensiva: remove espaços nas bordas dos nomes das colunas e dos IDs.
    df.columns = df.columns.astype(str).str.strip()
    for coluna in colunas_id:
        if coluna in df.columns:
            df[coluna] = df[coluna].str.strip().replace("", pd.NA)
    return df


def ler_clientes(pasta: Path) -> pd.DataFrame:
    """Lê Clientes.xlsx."""
    return _ler_planilha(pasta / ARQUIVO_CLIENTES, ["ID Cliente"])


def ler_estoque(pasta: Path) -> pd.DataFrame:
    """Lê Estoque.xlsx."""
    return _ler_planilha(pasta / ARQUIVO_ESTOQUE, ["ID Produto"])


def ler_pedidos(pasta: Path) -> pd.DataFrame:
    """Lê Pedidos.xlsx."""
    return _ler_planilha(
        pasta / ARQUIVO_PEDIDOS, ["ID Pedido", "ID Cliente", "ID Produto"]
    )