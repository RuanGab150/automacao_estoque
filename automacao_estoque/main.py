import customtkinter as ctk
from tkinter import ttk, messagebox
from excel import gerar_planilha


# ==========================================
# CONFIGURAÇÃO
# ==========================================

ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")


# ==========================================
# JANELA PRINCIPAL
# ==========================================

app = ctk.CTk()

app.title("Sistema de Entrada de Estoque")
app.geometry("1000x650")
app.minsize(900, 600)


# ==========================================
# TÍTULO
# ==========================================

titulo = ctk.CTkLabel(
    app,
    text="Sistema de Entrada de Estoque",
    font=("Arial", 30, "bold")
)

titulo.pack(pady=(50, 10))


# ==========================================
# SUBTÍTULO
# ==========================================

subtitulo = ctk.CTkLabel(
    app,
    text="Cadastro de novos pedidos",
    font=("Arial", 16)
)

subtitulo.pack(pady=(0, 40))


# ==========================================
# FUNÇÃO DO BOTÃO
# ==========================================

# ==========================================
# FUNÇÃO NOVA ENTRADA
# ==========================================

def abrir_nova_entrada():

    produtos = []

    # ==========================================
    # JANELA
    # ==========================================

    janela = ctk.CTkToplevel(app)

    janela.title("Nova Entrada de Produtos")
    janela.geometry("850x650")
    janela.minsize(800, 600)

    janela.focus()


    # ==========================================
    # TÍTULO
    # ==========================================

    titulo = ctk.CTkLabel(
        janela,
        text="Cadastro de Novo Produto",
        font=("Arial", 26, "bold")
    )

    titulo.pack(pady=(25, 20))


    # ==========================================
    # ÁREA DOS CAMPOS
    # ==========================================

    frame_campos = ctk.CTkFrame(janela)

    frame_campos.pack(
        padx=40,
        pady=10,
        fill="x"
    )


    # Faz as duas colunas ocuparem espaço
    frame_campos.grid_columnconfigure(0, weight=1)
    frame_campos.grid_columnconfigure(1, weight=1)


    # ==========================================
    # 1 - NOME DO PRODUTO
    # ==========================================

    label_nome = ctk.CTkLabel(
        frame_campos,
        text="Nome do Produto *",
        font=("Arial", 14, "bold")
    )

    label_nome.grid(
        row=0,
        column=0,
        padx=20,
        pady=(25, 5),
        sticky="w"
    )

    entrada_nome = ctk.CTkEntry(
        frame_campos,
        height=40,
        placeholder_text="Digite o nome do produto"
    )

    entrada_nome.grid(
        row=1,
        column=0,
        padx=20,
        pady=(0, 20),
        sticky="ew"
    )


    # ==========================================
    # 2 - GRADE FISCAL
    # ==========================================

    label_grade_fiscal = ctk.CTkLabel(
        frame_campos,
        text="Grade Fiscal *",
        font=("Arial", 14, "bold")
    )

    label_grade_fiscal.grid(
        row=0,
        column=1,
        padx=20,
        pady=(25, 5),
        sticky="w"
    )

    entrada_grade_fiscal = ctk.CTkEntry(
        frame_campos,
        height=40,
        placeholder_text="Digite a grade fiscal"
    )

    entrada_grade_fiscal.grid(
        row=1,
        column=1,
        padx=20,
        pady=(0, 20),
        sticky="ew"
    )


    # ==========================================
    # 3 - NCM
    # ==========================================

    label_ncm = ctk.CTkLabel(
        frame_campos,
        text="NCM *",
        font=("Arial", 14, "bold")
    )

    label_ncm.grid(
        row=2,
        column=0,
        padx=20,
        pady=(5, 5),
        sticky="w"
    )

    entrada_ncm = ctk.CTkEntry(
        frame_campos,
        height=40,
        placeholder_text="Ex: 8413.70.10"
    )

    entrada_ncm.grid(
        row=3,
        column=0,
        padx=20,
        pady=(0, 20),
        sticky="ew"
    )


    # ==========================================
    # 4 - SKU
    # ==========================================

    label_sku = ctk.CTkLabel(
        frame_campos,
        text="SKU *",
        font=("Arial", 14, "bold")
    )

    label_sku.grid(
        row=2,
        column=1,
        padx=20,
        pady=(5, 5),
        sticky="w"
    )

    entrada_sku = ctk.CTkEntry(
        frame_campos,
        height=40,
        placeholder_text="Digite o SKU"
    )

    entrada_sku.grid(
        row=3,
        column=1,
        padx=20,
        pady=(0, 20),
        sticky="ew"
    )


    # ==========================================
    # 5 - UNIDADE DE MEDIDA
    # ==========================================

    label_unidade = ctk.CTkLabel(
        frame_campos,
        text="Unidade de Medida *",
        font=("Arial", 14, "bold")
    )

    label_unidade.grid(
        row=4,
        column=0,
        padx=20,
        pady=(5, 5),
        sticky="w"
    )

    unidade_var = ctk.StringVar(
        value="Unidade"
    )

    entrada_unidade = ctk.CTkComboBox(
        frame_campos,
        height=40,
        variable=unidade_var,
        values=[
            "Unidade",
            "METRO",
            "PEÇA",
            "Ferramenta",
            "Pacote",
            "PAR"
        ]
    )

    entrada_unidade.grid(
        row=5,
        column=0,
        padx=20,
        pady=(0, 20),
        sticky="ew"
    )


    # ==========================================
    # 6 - PREÇO DE VENDA
    # ==========================================

    label_preco = ctk.CTkLabel(
        frame_campos,
        text="Preço de Venda *",
        font=("Arial", 14, "bold")
    )

    label_preco.grid(
        row=4,
        column=1,
        padx=20,
        pady=(5, 5),
        sticky="w"
    )

    entrada_preco = ctk.CTkEntry(
        frame_campos,
        height=40,
        placeholder_text="Ex: 49,90"
    )

    entrada_preco.grid(
        row=5,
        column=1,
        padx=20,
        pady=(0, 20),
        sticky="ew"
    )


    # ==========================================
    # 7 - ESTOQUE INICIAL
    # ==========================================

    label_estoque = ctk.CTkLabel(
        frame_campos,
        text="Estoque Inicial *",
        font=("Arial", 14, "bold")
    )

    label_estoque.grid(
        row=6,
        column=0,
        padx=20,
        pady=(5, 5),
        sticky="w"
    )

    entrada_estoque = ctk.CTkEntry(
        frame_campos,
        height=40,
        placeholder_text="Ex: 50"
    )

    entrada_estoque.grid(
        row=7,
        column=0,
        padx=20,
        pady=(0, 20),
        sticky="ew"
    )


    # ==========================================
    # FUNÇÃO ADICIONAR PRODUTO
    # ==========================================

    def adicionar_produto():

        # ==========================================
        # PEGA OS VALORES DOS CAMPOS
        # ==========================================

        nome = entrada_nome.get().strip()
        grade_fiscal = entrada_grade_fiscal.get().strip()
        ncm = entrada_ncm.get().strip()
        sku = entrada_sku.get().strip()
        unidade = unidade_var.get().strip()
        preco = entrada_preco.get().strip()
        estoque = entrada_estoque.get().strip()


        # ==========================================
        # VALIDA CAMPOS OBRIGATÓRIOS
        # ==========================================

        if not nome or not grade_fiscal or not ncm or not sku or not unidade or not preco or not estoque:

            messagebox.showwarning(
                title="Atenção",
                message="Preencha todos os campos obrigatórios antes de continuar."
            )

            return


        # ==========================================
        # VALIDA PREÇO
        # ==========================================

        try:

            preco_numerico = float(
                preco.replace(",", ".")
            )

        except ValueError:

            messagebox.showwarning(
                title="Preço inválido",
                message="O preço de venda deve ser um número válido.\n\nExemplo: 67,69"
            )

            entrada_preco.focus()

            return


        # ==========================================
        # VALIDA ESTOQUE
        # ==========================================

        try:

            estoque_numerico = float(
                estoque.replace(",", ".")
            )

        except ValueError:

            messagebox.showwarning(
                title="Estoque inválido",
                message="O estoque inicial deve ser um número válido.\n\nExemplo: 420"
            )

            entrada_estoque.focus()

            return


        # ==========================================
        # CRIA O PRODUTO
        # ==========================================

        produto = {

            "nome": nome,

            "grade_fiscal": grade_fiscal,

            "ncm": ncm,

            "sku": sku,

            "unidade": unidade,

            "preco_venda": preco_numerico,

            "estoque_inicial": estoque_numerico
        }


        # ==========================================
        # ADICIONA À LISTA
        # ==========================================

        produtos.append(produto)


        print("Produto adicionado:")
        print(produto)


        # ==========================================
        # ADICIONA À TABELA
        # ==========================================

        tabela.insert(
            "",
            "end",
            values=(
                produto["sku"],
                produto["nome"],
                produto["unidade"],
                f"{produto['preco_venda']:.2f}",
                produto["estoque_inicial"]
            )
        )


        # ==========================================
        # LIMPA OS CAMPOS
        # ==========================================

        entrada_nome.delete(0, "end")
        entrada_grade_fiscal.delete(0, "end")
        entrada_ncm.delete(0, "end")
        entrada_sku.delete(0, "end")
        entrada_preco.delete(0, "end")
        entrada_estoque.delete(0, "end")


        # Volta para o primeiro campo

        entrada_nome.focus()


    # ==========================================
    # LISTA DE PRODUTOS ADICIONADOS
    # ==========================================

    label_lista = ctk.CTkLabel(
        janela,
        text="Produtos Adicionados",
        font=("Arial", 16, "bold")
    )

    label_lista.pack(
        pady=(10, 5)
    )


    # ==========================================
    # TABELA
    # ==========================================

    frame_lista = ctk.CTkFrame(
        janela
    )

    frame_lista.pack(
    padx=40,
    pady=(0, 10),
    fill="x"
)

    frame_lista.configure(height=70)
    frame_lista.pack_propagate(False)


    tabela = ttk.Treeview(
        frame_lista,
        columns=(
            "sku",
            "nome",
            "unidade",
            "preco",
            "estoque"
        ),
        show="headings"
    )


    # ==========================================
    # CABEÇALHOS
    # ==========================================

    tabela.heading(
        "sku",
        text="SKU"
    )

    tabela.heading(
        "nome",
        text="Produto"
    )

    tabela.heading(
        "unidade",
        text="Unidade"
    )

    tabela.heading(
        "preco",
        text="Preço"
    )

    tabela.heading(
        "estoque",
        text="Estoque"
    )


    # ==========================================
    # LARGURA DAS COLUNAS
    # ==========================================

    tabela.column(
        "sku",
        width=120
    )

    tabela.column(
        "nome",
        width=300
    )

    tabela.column(
        "unidade",
        width=120
    )

    tabela.column(
        "preco",
        width=100
    )

    tabela.column(
        "estoque",
        width=100
    )


    tabela.pack(
        fill="both",
        expand=True,
        padx=5,
        pady=5
    )

    # ==========================================
    # FUNÇÃO TERMINAR PLANILHA
    # ==========================================

    def terminar_planilha():

        # ==========================================
        # VERIFICA SE EXISTE PELO MENOS UM PRODUTO
        # ==========================================

        if not produtos:

            messagebox.showwarning(
                title="Nenhum produto",
                message="Adicione pelo menos um produto antes de terminar a planilha."
            )

            return


        # ==========================================
        # GERA A PLANILHA
        # ==========================================

        try:

            gerar_planilha(produtos)


            # ======================================
            # AVISA QUE DEU CERTO
            # ======================================

            messagebox.showinfo(
                title="Sucesso",
                message=(
                    f"Planilha gerada com sucesso!\n\n"
                    f"Produtos adicionados: {len(produtos)}\n\n"
                    f"Arquivo salvo na pasta 'output'."
                )
            )


            # Fecha a janela de cadastro

            janela.destroy()


        except Exception as erro:

            messagebox.showerror(
                title="Erro",
                message=f"Não foi possível gerar a planilha.\n\n{erro}"
            )


    # ==========================================
    # ÁREA DOS BOTÕES
    # ==========================================

    frame_botoes = ctk.CTkFrame(
        janela
    )

    frame_botoes.pack(
        pady=(5, 20)
    )


    # ==========================================
    # BOTÃO ADICIONAR PRODUTO
    # ==========================================

    botao_adicionar = ctk.CTkButton(
        frame_botoes,
        text="Adicionar Mais um Produto",
        width=280,
        height=45,
        command=adicionar_produto
    )

    botao_adicionar.grid(
        row=0,
        column=0,
        padx=10
    )


    # ==========================================
    # BOTÃO TERMINAR PLANILHA
    # ==========================================

    botao_terminar = ctk.CTkButton(
        frame_botoes,
        text="Terminar Planilha",
        width=280,
        height=45,
        command=terminar_planilha
    )

    botao_terminar.grid(
        row=0,
        column=1,
        padx=10
    )

# ==========================================
# BOTÃO NOVA ENTRADA
# ==========================================

botao_nova_entrada = ctk.CTkButton(
    app,
    text="Nova Entrada",
    width=200,
    height=40,
    command=abrir_nova_entrada
)

botao_nova_entrada.pack(pady=10)


# ==========================================
# BOTÃO SAIR
# ==========================================

botao_sair = ctk.CTkButton(
    app,
    text="Sair",
    width=200,
    height=40,
    command=app.destroy
)

botao_sair.pack(pady=10)


# ==========================================
# INICIA O PROGRAMA
# ==========================================

app.mainloop()