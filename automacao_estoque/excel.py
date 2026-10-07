import openpyxl
import os
import sys


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

        aba[f"I{linha}"] = produto["grade_fiscal"]

        aba[f"J{linha}"] = produto["ncm"]

        aba[f"P{linha}"] = produto["sku"]

        aba[f"T{linha}"] = produto["unidade"]

        aba[f"V{linha}"] = produto["preco_venda"]

        aba[f"AI{linha}"] = produto["estoque_inicial"]

        linha += 1


    # ==========================================
    # CRIA A PASTA OUTPUT
    # ==========================================

    pasta_saida = os.path.join(
        pasta_programa,
        "output"
    )

    os.makedirs(
        pasta_saida,
        exist_ok=True
    )


    # ==========================================
    # NOME DO ARQUIVO
    # ==========================================

    arquivo_saida = os.path.join(
        pasta_saida,
        "pedido.xlsx"
    )


    # ==========================================
    # SALVA A PLANILHA
    # ==========================================

    planilha.save(
        arquivo_saida
    )


    print("===================================")
    print("PLANILHA GERADA COM SUCESSO!")
    print("===================================")
    print(f"Produtos adicionados: {len(produtos)}")
    print(f"Arquivo: {arquivo_saida}")