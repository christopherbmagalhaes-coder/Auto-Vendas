"""Validação dos dados de entrada (erros impeditivos x avisos)."""
from dataclasses import dataclass, field

import pandas as pd

COLUNAS_OBRIGATORIAS = {
    "Clientes": ["ID Cliente", "Cliente"],
    "Estoque": ["ID Produto", "Produto", "Preço unit.", "Quantidade"],
    "Pedidos": ["ID Pedido", "Data", "ID Cliente", "ID Produto", "Qtd.", "Total"],
}
LIMITE_EXIBIDO = 10
TOLERANCIA_TOTAL = 0.01


@dataclass
class ResultadoValidacao:
    """Agrupa erros (impeditivos) e avisos (informativos)."""

    erros: list[str] = field(default_factory=list)
    avisos: list[str] = field(default_factory=list)

    @property
    def tem_erros(self) -> bool:
        return bool(self.erros)

    def adicionar(self, outro: "ResultadoValidacao") -> None:
        self.erros.extend(outro.erros)
        self.avisos.extend(outro.avisos)


# =-=-=-=-=-=-= Utilitários (também usados em processamento.py) =-=-=-=-=-=-=

def converter_numero(serie: pd.Series) -> pd.Series:
    """Converte para número; valores inválidos viram NaN."""
    return pd.to_numeric(serie, errors="coerce")


def converter_data(serie: pd.Series) -> pd.Series:
    """Converte para data (dia primeiro); valores inválidos viram NaT.

    Números soltos (ex.: 45000) são tratados como inválidos para evitar
    interpretações silenciosamente erradas.
    """
    if pd.api.types.is_numeric_dtype(serie):
        return pd.Series(pd.NaT, index=serie.index, dtype="datetime64[ns]")
    datas = pd.to_datetime(serie, errors="coerce", dayfirst=True, format="mixed")
    numericos = serie.map(lambda valor: isinstance(valor, (int, float)))
    return datas.mask(numericos)


def _linhas(mascara: pd.Series) -> str:
    """Lista os números de linha do Excel (cabeçalho = linha 1) onde a máscara é verdadeira."""
    linhas = [int(indice) + 2 for indice in mascara[mascara].index]
    texto = ", ".join(str(linha) for linha in linhas[:LIMITE_EXIBIDO])
    if len(linhas) > LIMITE_EXIBIDO:
        texto += f" (e mais {len(linhas) - LIMITE_EXIBIDO})"
    return texto


def _vazio(serie: pd.Series) -> pd.Series:
    return serie.isna() | serie.astype(str).str.strip().eq("")


# =-=-=-=-=-=-= Validações de estrutura =-=-=-=-=-=-=

def validar_colunas(
    clientes: pd.DataFrame, estoque: pd.DataFrame, pedidos: pd.DataFrame
) -> ResultadoValidacao:
    """Confirma a existência das colunas obrigatórias em cada arquivo."""
    resultado = ResultadoValidacao()
    tabelas = {"Clientes": clientes, "Estoque": estoque, "Pedidos": pedidos}
    for nome, df in tabelas.items():
        ausentes = [c for c in COLUNAS_OBRIGATORIAS[nome] if c not in df.columns]
        if ausentes:
            resultado.erros.append(
                f"{nome}.xlsx: coluna(s) obrigatória(s) ausente(s): {', '.join(ausentes)}. "
                f"Colunas encontradas: {', '.join(df.columns) or '(nenhuma)'}."
            )
    return resultado


# =-=-=-=-=-= Validação dos cadastros =-=-=-=-=-=

def _validar_id_unico(resultado, df, coluna, arquivo):
    vazios = df[coluna].isna()
    if vazios.any():
        resultado.erros.append(f"{arquivo}: '{coluna}' vazio nas linhas {_linhas(vazios)}.")

    duplicados = df[coluna].notna() & df.duplicated(coluna, keep=False)
    if duplicados.any():
        ids = sorted(df.loc[duplicados, coluna].unique())[:LIMITE_EXIBIDO]
        resultado.erros.append(
            f"{arquivo}: '{coluna}' duplicado ({', '.join(ids)}) nas linhas {_linhas(duplicados)}."
        )


def _validar_texto_obrigatorio(resultado, df, coluna, arquivo):
    vazios = _vazio(df[coluna])
    if vazios.any():
        resultado.erros.append(f"{arquivo}: '{coluna}' vazio nas linhas {_linhas(vazios)}.")


def validar_cadastros(clientes: pd.DataFrame, estoque: pd.DataFrame) -> ResultadoValidacao:
    """Valida unicidade de IDs, nomes e valores do estoque."""
    resultado = ResultadoValidacao()

    _validar_id_unico(resultado, clientes, "ID Cliente", "Clientes")
    _validar_texto_obrigatorio(resultado, clientes, "Cliente", "Clientes")

    _validar_id_unico(resultado, estoque, "ID Produto", "Estoque")
    _validar_texto_obrigatorio(resultado, estoque, "Produto", "Estoque")

    preco = converter_numero(estoque["Preço unit."])
    invalido = preco.isna() | (preco < 0)
    if invalido.any():
        resultado.erros.append(
            f"Estoque: 'Preço unit.' ausente, não numérico ou negativo nas linhas {_linhas(invalido)}."
        )

    # Quantidade não entra no relatório nesta versão: problemas geram apenas aviso.
    quantidade = converter_numero(estoque["Quantidade"])
    invalida = quantidade.isna() | (quantidade < 0)
    if invalida.any():
        resultado.avisos.append(
            f"Estoque: 'Quantidade' ausente, não numérica ou negativa nas linhas {_linhas(invalida)} "
            "(não afeta o relatório de vendas)."
        )
    return resultado


# =-=-=-=-=-=-= Validação dos pedidos =-=-=-=-=-=-=

def _validar_referencias(resultado, pedidos, cadastro, coluna, descricao):
    existentes = cadastro[coluna].dropna()
    inexistentes = pedidos[coluna].notna() & ~pedidos[coluna].isin(existentes)
    if inexistentes.any():
        ids = sorted(pedidos.loc[inexistentes, coluna].unique())[:LIMITE_EXIBIDO]
        resultado.erros.append(
            f"Pedidos: '{coluna}' sem correspondência em {descricao} "
            f"({', '.join(ids)}) nas linhas {_linhas(inexistentes)}."
        )


def validar_pedidos(
    pedidos: pd.DataFrame, clientes: pd.DataFrame, estoque: pd.DataFrame
) -> ResultadoValidacao:
    """Valida campos, tipos, referências e coerência dos pedidos."""
    resultado = ResultadoValidacao()

    if pedidos.empty:
        resultado.avisos.append(
            "Pedidos.xlsx não possui registros; a planilha de vendas será gerada vazia."
        )
        return resultado

    for coluna in ("ID Pedido", "ID Cliente", "ID Produto"):
        vazios = pedidos[coluna].isna()
        if vazios.any():
            resultado.erros.append(f"Pedidos: '{coluna}' vazio nas linhas {_linhas(vazios)}.")

    _validar_referencias(resultado, pedidos, clientes, "ID Cliente", "Clientes")
    _validar_referencias(resultado, pedidos, estoque, "ID Produto", "Estoque")

    datas_invalidas = converter_data(pedidos["Data"]).isna()
    if datas_invalidas.any():
        resultado.erros.append(
            f"Pedidos: 'Data' ausente ou inválida nas linhas {_linhas(datas_invalidas)}."
        )

    qtd = converter_numero(pedidos["Qtd."])
    qtd_invalida = qtd.isna() | (qtd <= 0) | (qtd % 1 != 0)
    if qtd_invalida.any():
        resultado.erros.append(
            f"Pedidos: 'Qtd.' deve ser um inteiro positivo; problema nas linhas {_linhas(qtd_invalida)}."
        )

    total = converter_numero(pedidos["Total"])
    total_invalido = total.isna() | (total < 0)
    if total_invalido.any():
        resultado.erros.append(
            f"Pedidos: 'Total' ausente, não numérico ou negativo nas linhas {_linhas(total_invalido)}."
        )

    # =-=-=-=-=-=-= Avisos (não impedem a geração) =-=-=-=-=-=-=
    repetidos = pedidos.duplicated(["ID Pedido", "ID Produto"], keep=False)
    if repetidos.any():
        resultado.avisos.append(
            "Pedidos: mesmo 'ID Pedido' + 'ID Produto' em mais de uma linha "
            f"(linhas {_linhas(repetidos)}). Todas foram mantidas."
        )

    cadastro_unico = estoque.dropna(subset=["ID Produto"]).drop_duplicates("ID Produto")
    precos = converter_numero(cadastro_unico.set_index("ID Produto")["Preço unit."])
    esperado = pedidos["ID Produto"].map(precos) * qtd
    divergente = (
        esperado.notna() & total.notna() & ((total - esperado).abs() > TOLERANCIA_TOTAL)
    )
    if divergente.any():
        resultado.avisos.append(
            "Pedidos: 'Total' difere de Qtd. × preço atual do estoque nas linhas "
            f"{_linhas(divergente)}. Os totais originais foram preservados "
            "(o preço pode ter mudado desde a venda)."
        )
    return resultado