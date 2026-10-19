"""Cruzamento e organização dos dados de vendas."""
import pandas as pd

from validacao import converter_data, converter_numero

COLUNAS_FINAIS = ["ID", "Pedido", "Data", "Cliente", "Produto", "Qtd.", "Preço unit.", "Total"]


class ErroDeProcessamento(Exception):
    """Falha ao relacionar ou organizar os dados."""


def _cruzar(pedidos, clientes, estoque) -> pd.DataFrame:
    """Faz os dois merges garantindo que não haja multiplicação de linhas."""
    base = pedidos[["ID Pedido", "Data", "ID Cliente", "ID Produto", "Qtd.", "Total"]]
    try:
        vendas = base.merge(
            clientes[["ID Cliente", "Cliente"]],
            on="ID Cliente", how="left", validate="many_to_one",
        )
        vendas = vendas.merge(
            estoque[["ID Produto", "Produto", "Preço unit."]],
            on="ID Produto", how="left", validate="many_to_one",
        )
    except pd.errors.MergeError as erro:
        raise ErroDeProcessamento(f"IDs duplicados impediram o cruzamento: {erro}") from erro
    return vendas


def _garantir_referencias(vendas: pd.DataFrame) -> None:
    """Rede de segurança: nenhum pedido pode ficar sem cliente ou produto."""
    sem_referencia = vendas["Cliente"].isna() | vendas["Produto"].isna()
    if sem_referencia.any():
        raise ErroDeProcessamento(
            f"{int(sem_referencia.sum())} pedido(s) sem cliente ou produto correspondente."
        )


def gerar_vendas(
    clientes: pd.DataFrame, estoque: pd.DataFrame, pedidos: pd.DataFrame
) -> pd.DataFrame:
    """Gera o DataFrame final de vendas, pronto para exportação."""
    vendas = _cruzar(pedidos, clientes, estoque)
    _garantir_referencias(vendas)

    vendas["Data"] = converter_data(vendas["Data"]).dt.normalize()
    vendas["Qtd."] = converter_numero(vendas["Qtd."]).astype(int)
    vendas["Preço unit."] = converter_numero(vendas["Preço unit."]).astype(float)
    vendas["Total"] = converter_numero(vendas["Total"]).astype(float)  # valor original preservado

    vendas = vendas.rename(columns={"ID Pedido": "Pedido"})
    vendas.insert(0, "ID", range(1, len(vendas) + 1))
    return vendas[COLUNAS_FINAIS].reset_index(drop=True)