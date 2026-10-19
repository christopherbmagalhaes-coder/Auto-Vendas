import sys
from pathlib import Path

from exportacao import ErroDeExportacao, exportar_vendas
from formatacao import ErroDeFormatacao, formatar_planilha
from leitura import ErroDeLeitura, ler_clientes, ler_estoque, ler_pedidos
from processamento import ErroDeProcessamento, gerar_vendas
from validacao import (
    ResultadoValidacao,
    validar_cadastros,
    validar_colunas,
    validar_pedidos,
)

PASTA_BASE = Path(__file__).resolve().parent
CAMINHO_SAIDA = PASTA_BASE / "Vendas.xlsx"


def ler_entradas():
    """Lê os três arquivos, acumulando os erros de leitura de todos eles."""
    leitores = {
        "clientes": ler_clientes,
        "estoque": ler_estoque,
        "pedidos": ler_pedidos,
    }
    dados, erros = {}, []
    for nome, leitor in leitores.items():
        try:
            dados[nome] = leitor(PASTA_BASE)
        except ErroDeLeitura as erro:
            erros.append(str(erro))
    return dados, erros


def exibir(titulo, mensagens):
    """Imprime uma lista de mensagens com os títulos."""
    if mensagens:
        print(f"\n{titulo} ({len(mensagens)}):")
        for mensagem in mensagens:
            print(f"  - {mensagem}")


def validar_tudo(clientes, estoque, pedidos) -> ResultadoValidacao:
    """Executa as validações; só avança se as colunas obrigatórias existirem."""
    resultado = validar_colunas(clientes, estoque, pedidos)
    if not resultado.tem_erros:
        resultado.adicionar(validar_cadastros(clientes, estoque))
        resultado.adicionar(validar_pedidos(pedidos, clientes, estoque))
    return resultado


def main() -> int:
    """Coordena leitura, validação, processamento, exportação e formatação."""
    print("Iniciando automação de vendas...")

    dados, erros_leitura = ler_entradas()
    if erros_leitura:
        exibir("ERROS DE LEITURA", erros_leitura)
        print("\nProcessamento interrompido. Nenhum arquivo foi gerado.")
        return 1

    clientes, estoque, pedidos = dados["clientes"], dados["estoque"], dados["pedidos"]

    resultado = validar_tudo(clientes, estoque, pedidos)
    exibir("AVISOS", resultado.avisos)
    exibir("ERROS IMPEDITIVOS", resultado.erros)
    if resultado.tem_erros:
        print(
            "\nProcessamento cancelado: verifique os erros acima. "
            "Nenhum arquivo foi criado ou alterado :( "
        )
        return 1

    try:
        vendas = gerar_vendas(clientes, estoque, pedidos)
        exportar_vendas(vendas, CAMINHO_SAIDA)
        formatar_planilha(CAMINHO_SAIDA)
    except (ErroDeProcessamento, ErroDeExportacao, ErroDeFormatacao) as erro:
        print(f"\nERRO: {erro}")
        return 1

    print(f"\nSucesso!!! {len(vendas)} registro(s) exportado(s) para '{CAMINHO_SAIDA.name}'.")
    return 0


if __name__ == "__main__":
    sys.exit(main())