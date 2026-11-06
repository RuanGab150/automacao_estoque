import openpyxl
import os
import sys
from datetime import datetime


def gerar_planilha(produtos):

    # ==========================================
    # LOCAL DO PROGRAMA
    # ==========================================

    if getattr(sys, "frozen", False):

        # Quando estiver rodando como .exe
        pasta_programa = os.path.dirname(
            sys.executable
        )

    else:

        # Quando estiver rodando pelo Python
        pasta_programa = os.path.dirname(
            os.path.abspath(__file__)
        )


    # ==========================================
    # ABRE O MODELO
    # ==========================================

    arquivo_modelo = os.path.join(
        pasta_programa,
        "Exemplo.xlsx"
    )

    if not os.path.exists(arquivo_modelo):

        raise FileNotFoundError(
            f"Modelo 'Exemplo.xlsx' não encontrado em:\n{arquivo_modelo}\n\n"
            f"Ele precisa ficar na mesma pasta do programa."
        )

    planilha = openpyxl.load_workbook(
        arquivo_modelo
    )

    aba = planilha["Dados"]


    # ==========================================
    # COMEÇA NA LINHA 3
    # ==========================================

    linha = 3


    # ==========================================
    # PERCORRE TODOS OS PRODUTOS
    # ==========================================

    for produto in produtos:

        aba[f"E{linha}"] = produto["nome"]

        # Marca (opcional) - coluna H
        aba[f"H{linha}"] = produto.get("marca", "")

        aba[f"I{linha}"] = produto["grade_fiscal"]

        aba[f"J{linha}"] = produto["ncm"]

        aba[f"P{linha}"] = produto["sku"]


        # ==========================================
        # NOVO: CÓDIGO DE BARRAS (COLUNA S)
        # Gravado como texto para o Excel não usar
        # notação científica nem cortar zeros
        # ==========================================

        codigo_barras = produto.get("codigo_barras", "")

        if codigo_barras:

            aba[f"S{linha}"] = str(codigo_barras)

            aba[f"S{linha}"].number_format = "@"


        aba[f"T{linha}"] = produto["unidade"]

        aba[f"V{linha}"] = produto["preco_venda"]

        aba[f"AI{linha}"] = produto["estoque_inicial"]

        linha += 1


    # ==========================================
    # NOME DO ARQUIVO (com data e hora)
    # ==========================================
    # Nome único: nunca sobrescreve a planilha anterior e
    # não dá conflito se o pedido.xlsx antigo estiver aberto.

    nome_arquivo = datetime.now().strftime(
        "pedido_%Y-%m-%d_%H-%M-%S.xlsx"
    )


    # ==========================================
    # PASTA OUTPUT (ao lado do programa)
    # ==========================================
    # Se não for possível gravar ao lado do programa
    # (pasta protegida, rede, somente leitura), usa
    # Documentos\SistemaEstoque\output como alternativa.

    candidatas = [
        os.path.join(pasta_programa, "output"),
        os.path.join(
            os.path.expanduser("~"),
            "Documents",
            "SistemaEstoque",
            "output"
        ),
    ]

    arquivo_saida = None
    ultimo_erro = None

    for pasta_saida in candidatas:

        try:

            os.makedirs(
                pasta_saida,
                exist_ok=True
            )

            caminho = os.path.abspath(
                os.path.join(pasta_saida, nome_arquivo)
            )

            planilha.save(
                caminho
            )

            # Confirma que o arquivo realmente existe no disco
            if os.path.exists(caminho):
                arquivo_saida = caminho
                break

        except OSError as erro:

            ultimo_erro = erro

    if arquivo_saida is None:

        raise OSError(
            f"Não foi possível salvar a planilha.\n\n{ultimo_erro}"
        )


    print("===================================")
    print("PLANILHA GERADA COM SUCESSO!")
    print("===================================")
    print(f"Produtos adicionados: {len(produtos)}")
    print(f"Arquivo: {arquivo_saida}")

    # Devolve o caminho completo para o programa mostrar ao usuário
    return arquivo_saida