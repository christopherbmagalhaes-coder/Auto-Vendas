"""Exportação do DataFrame de vendas para Excel."""
from pathlib import Path

import pandas as pd

NOME_ABA = "Vendas"


class ErroDeExportacao(Exception):
    """Falha ao gravar o arquivo de saída."""


def exportar_vendas(vendas: pd.DataFrame, caminho_saida: Path) -> None:
    """Grava o DataFrame em Excel, sem estilos e sem índice."""
    try:
        vendas.to_excel(caminho_saida, sheet_name=NOME_ABA, index=False, engine="openpyxl")
    except PermissionError as erro:
        raise ErroDeExportacao(
            f"Sem permissão para gravar '{Path(caminho_saida).name}'. "
            "Feche o arquivo se estiver aberto no Excel e tente novamente."
        ) from erro
    except OSError as erro:
        raise ErroDeExportacao(
            f"Não foi possível gravar '{Path(caminho_saida).name}': {erro}"
        ) from erro