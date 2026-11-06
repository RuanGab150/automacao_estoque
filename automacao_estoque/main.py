import customtkinter as ctk
from tkinter import ttk, messagebox
import os
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
# FUNÇÃO NOVA ENTRADA
# ==========================================

def abrir_nova_entrada():

    produtos = []

    # Liga cada linha da tabela ao produto correspondente
    produtos_por_linha = {}

    # Linha da tabela que está sendo editada (None = cadastrando novo)
    estado = {"editando": None}


    # ==========================================
    # JANELA
    # ==========================================

    janela = ctk.CTkToplevel(app)

    janela.title("Nova Entrada de Produtos")
    janela.minsize(1000, 650)

    # A janela NÃO usa tela cheia: ocupa a área útil do Windows
    # (a resolução menos a barra de tarefas), então nada fica escondido.
    # O tamanho é calculado depois que a janela existe, para o Windows
    # informar a área útil já descontando a barra de tarefas.
    largura_tela = janela.winfo_screenwidth()
    altura_tela = janela.winfo_screenheight()

    janela.geometry(f"{largura_tela}x{altura_tela}+0+0")

    # A janela fica ligada à principal (sempre na frente dela)
    janela.transient(app)


    # ==========================================
    # TRAZ PARA A FRENTE E MAXIMIZA
    # ==========================================
    # O CustomTkinter termina de montar o CTkToplevel uns 200 ms
    # depois de criado e desfaz a maximização. Por isso o atraso.

    def trazer_para_frente():

        # NÃO usa tela cheia: o Windows diz qual é a área útil da tela
        # (a resolução menos a barra de tarefas) e a janela ocupa só ela,
        # com uma folga. Assim nenhum botão fica atrás da barra de tarefas.
        janela.update_idletasks()

        margem = 12

        largura = janela.winfo_screenwidth() - (margem * 2)
        altura = janela.winfo_screenheight() - (margem * 3)

        janela.geometry(
            f"{largura}x{altura}+{margem}+{margem}"
        )

        janela.lift()
        janela.attributes("-topmost", True)
        janela.focus_force()

        # Tira o "sempre no topo" para não ficar na frente de outros programas
        janela.after(300, lambda: janela.attributes("-topmost", False))

        # Cursor já no campo certo para o leitor
        if usar_codigo_barras.get():
            entrada_codigo_barras.focus_set()
        else:
            entrada_nome.focus_set()

    janela.after(400, trazer_para_frente)


    # ==========================================
    # TÍTULO
    # ==========================================

    titulo = ctk.CTkLabel(
        janela,
        text="Cadastro de Novo Produto",
        font=("Arial", 26, "bold")
    )

    titulo.pack(pady=(15, 5))


    # ==========================================
    # ÁREA DOS CAMPOS (3 COLUNAS)
    # ==========================================

    frame_campos = ctk.CTkFrame(janela)

    frame_campos.pack(
        padx=40,
        pady=5,
        fill="x"
    )

    frame_campos.grid_columnconfigure(0, weight=1)
    frame_campos.grid_columnconfigure(1, weight=1)
    frame_campos.grid_columnconfigure(2, weight=1)


    def criar_label(texto, linha, coluna):

        label = ctk.CTkLabel(
            frame_campos,
            text=texto,
            font=("Arial", 14, "bold")
        )

        label.grid(
            row=linha,
            column=coluna,
            padx=20,
            pady=(10, 2),
            sticky="w"
        )

        return label


    def criar_entrada(placeholder, linha, coluna):

        entrada = ctk.CTkEntry(
            frame_campos,
            height=38,
            placeholder_text=placeholder
        )

        entrada.grid(
            row=linha,
            column=coluna,
            padx=20,
            pady=(0, 8),
            sticky="ew"
        )

        return entrada


    # ---------- LINHA 1 ----------

    # 1 - NOME DO PRODUTO
    criar_label("Nome do Produto *", 0, 0)
    entrada_nome = criar_entrada("Digite o nome do produto", 1, 0)

    # 2 - GRADE FISCAL
    criar_label("Grade Fiscal *", 0, 1)
    entrada_grade_fiscal = criar_entrada("Digite a grade fiscal", 1, 1)

    # 3 - NCM
    criar_label("NCM *", 0, 2)
    entrada_ncm = criar_entrada("Ex: 8413.70.10", 1, 2)


    # ---------- LINHA 2 ----------

    # 4 - SKU
    criar_label("SKU *", 2, 0)
    entrada_sku = criar_entrada("Digite o SKU", 3, 0)

    # 5 - NOVO: MARCA (OPCIONAL)
    criar_label("Marca", 2, 1)
    entrada_marca = criar_entrada("Digite a marca (opcional)", 3, 1)

    # 6 - UNIDADE DE MEDIDA
    criar_label("Unidade de Medida *", 2, 2)

    unidade_var = ctk.StringVar(
        value="Unidade"
    )

    entrada_unidade = ctk.CTkComboBox(
        frame_campos,
        height=38,
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
        row=3,
        column=2,
        padx=20,
        pady=(0, 8),
        sticky="ew"
    )


    # ---------- LINHA 3 ----------

    # 7 - PREÇO DE VENDA
    criar_label("Preço de Venda *", 4, 0)
    entrada_preco = criar_entrada("Ex: 49,90", 5, 0)

    # 8 - ESTOQUE INICIAL
    criar_label("Estoque Inicial *", 4, 1)
    entrada_estoque = criar_entrada("Ex: 50", 5, 1)

    # 9 - CÓDIGO DE BARRAS (OPCIONAL)
    label_codigo_barras = criar_label("Código de Barras", 4, 2)
    entrada_codigo_barras = criar_entrada("Digite ou leia com o leitor", 5, 2)


    # O leitor aperta Enter no final da leitura:
    # o código fica no campo e o cursor vai para o Nome.
    def leitura_concluida(evento):

        entrada_nome.focus_set()

        return "break"

    entrada_codigo_barras.bind("<Return>", leitura_concluida)
    entrada_codigo_barras.bind("<KP_Enter>", leitura_concluida)


    # ==========================================
    # OPÇÃO ATIVAR/DESATIVAR (ABAIXO DO CAMPO)
    # ==========================================

    usar_codigo_barras = ctk.BooleanVar(value=True)

    switch_codigo_barras = ctk.CTkSwitch(
        frame_campos,
        text="Usar código de barras",
        variable=usar_codigo_barras,
        font=("Arial", 14)
    )

    switch_codigo_barras.grid(
        row=6,
        column=2,
        padx=20,
        pady=(0, 12),
        sticky="w"
    )


    COLUNAS_COM_CODIGO = (
        "sku", "nome", "marca", "codigo_barras",
        "unidade", "preco", "estoque"
    )

    COLUNAS_SEM_CODIGO = (
        "sku", "nome", "marca",
        "unidade", "preco", "estoque"
    )


    def alternar_codigo_barras():

        if usar_codigo_barras.get():

            label_codigo_barras.grid()
            entrada_codigo_barras.grid()
            entrada_codigo_barras.focus_set()

            tabela.configure(displaycolumns=COLUNAS_COM_CODIGO)

        else:

            entrada_codigo_barras.delete(0, "end")
            label_codigo_barras.grid_remove()
            entrada_codigo_barras.grid_remove()

            tabela.configure(displaycolumns=COLUNAS_SEM_CODIGO)


    switch_codigo_barras.configure(
        command=alternar_codigo_barras
    )


    # ==========================================
    # FUNÇÕES AUXILIARES
    # ==========================================

    def limpar_campos():

        entrada_nome.delete(0, "end")
        entrada_grade_fiscal.delete(0, "end")
        entrada_ncm.delete(0, "end")
        entrada_sku.delete(0, "end")
        entrada_marca.delete(0, "end")
        entrada_codigo_barras.delete(0, "end")
        entrada_preco.delete(0, "end")
        entrada_estoque.delete(0, "end")
        unidade_var.set("Unidade")


    def focar_primeiro_campo():

        if usar_codigo_barras.get():
            entrada_codigo_barras.focus_set()
        else:
            entrada_nome.focus_set()


    def formatar_numero(valor):

        # 50.0 vira "50" e 49.9 vira "49,9" ao voltar para o campo
        if float(valor).is_integer():
            return str(int(valor))

        return str(valor).replace(".", ",")


    def valores_da_tabela(produto):

        return (
            produto["sku"],
            produto["nome"],
            produto["marca"],
            produto["codigo_barras"],
            produto["unidade"],
            f"{produto['preco_venda']:.2f}",
            produto["estoque_inicial"]
        )


    def sair_do_modo_edicao():

        estado["editando"] = None

        titulo.configure(text="Cadastro de Novo Produto")
        botao_adicionar.configure(text="Adicionar Mais um Produto")
        botao_cancelar.grid_remove()

        limpar_campos()
        focar_primeiro_campo()


    # ==========================================
    # FUNÇÃO ADICIONAR / SALVAR PRODUTO
    # ==========================================

    def adicionar_produto():

        # ==========================================
        # PEGA OS VALORES DOS CAMPOS
        # ==========================================

        nome = entrada_nome.get().strip()
        grade_fiscal = entrada_grade_fiscal.get().strip()
        ncm = entrada_ncm.get().strip()
        sku = entrada_sku.get().strip()
        marca = entrada_marca.get().strip()
        unidade = unidade_var.get().strip()
        preco = entrada_preco.get().strip()
        estoque = entrada_estoque.get().strip()

        codigo_barras = ""

        if usar_codigo_barras.get():
            codigo_barras = entrada_codigo_barras.get().strip()


        # ==========================================
        # VALIDA CAMPOS OBRIGATÓRIOS
        # ==========================================

        if not nome or not grade_fiscal or not ncm or not sku or not unidade or not preco or not estoque:

            messagebox.showwarning(
                title="Atenção",
                message="Preencha todos os campos obrigatórios antes de continuar.",
                parent=janela
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
                message="O preço de venda deve ser um número válido.\n\nExemplo: 67,69",
                parent=janela
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
                message="O estoque inicial deve ser um número válido.\n\nExemplo: 420",
                parent=janela
            )

            entrada_estoque.focus()

            return


        # ==========================================
        # VALIDA CÓDIGO DE BARRAS (OPCIONAL)
        # ==========================================

        if codigo_barras and not codigo_barras.isdigit():

            messagebox.showwarning(
                title="Código de barras inválido",
                message="O código de barras deve conter apenas números.",
                parent=janela
            )

            entrada_codigo_barras.focus()

            return


        # ==========================================
        # MONTA O PRODUTO
        # ==========================================

        produto = {

            "nome": nome,

            "grade_fiscal": grade_fiscal,

            "ncm": ncm,

            "sku": sku,

            "marca": marca,

            "codigo_barras": codigo_barras,

            "unidade": unidade,

            "preco_venda": preco_numerico,

            "estoque_inicial": estoque_numerico
        }


        # ==========================================
        # MODO EDIÇÃO: ATUALIZA O PRODUTO EXISTENTE
        # ==========================================

        linha_editada = estado["editando"]

        if linha_editada is not None:

            # Atualiza o mesmo dicionário que está na lista "produtos"
            produtos_por_linha[linha_editada].update(produto)

            tabela.item(
                linha_editada,
                values=valores_da_tabela(produtos_por_linha[linha_editada])
            )

            print("Produto alterado:")
            print(produtos_por_linha[linha_editada])

            sair_do_modo_edicao()

            return


        # ==========================================
        # MODO NORMAL: ADICIONA NOVO PRODUTO
        # ==========================================

        produtos.append(produto)

        linha = tabela.insert(
            "",
            "end",
            values=valores_da_tabela(produto)
        )

        produtos_por_linha[linha] = produto

        print("Produto adicionado:")
        print(produto)

        limpar_campos()
        focar_primeiro_campo()


    # ==========================================
    # FUNÇÃO EDITAR PRODUTO SELECIONADO
    # ==========================================

    def editar_selecionado(evento=None):

        selecao = tabela.selection()

        if not selecao:

            messagebox.showwarning(
                title="Nenhum produto selecionado",
                message="Clique em um produto da lista para editar.",
                parent=janela
            )

            return

        linha = selecao[0]
        produto = produtos_por_linha[linha]

        limpar_campos()

        entrada_nome.insert(0, produto["nome"])
        entrada_grade_fiscal.insert(0, produto["grade_fiscal"])
        entrada_ncm.insert(0, produto["ncm"])
        entrada_sku.insert(0, produto["sku"])
        entrada_marca.insert(0, produto["marca"])
        unidade_var.set(produto["unidade"])
        entrada_preco.insert(0, formatar_numero(produto["preco_venda"]))
        entrada_estoque.insert(0, formatar_numero(produto["estoque_inicial"]))

        # Se o produto tem código de barras, liga o campo para mostrar
        if produto["codigo_barras"] and not usar_codigo_barras.get():
            usar_codigo_barras.set(True)
            alternar_codigo_barras()

        entrada_codigo_barras.insert(0, produto["codigo_barras"])

        estado["editando"] = linha

        titulo.configure(text="Editando Produto")
        botao_adicionar.configure(text="Salvar Alteração")
        botao_cancelar.grid()

        entrada_nome.focus_set()

        # Evita que o duplo clique faça outra coisa na tabela
        return "break"


    # ==========================================
    # FUNÇÃO REMOVER PRODUTO SELECIONADO
    # ==========================================

    def remover_selecionado():

        selecao = tabela.selection()

        if not selecao:

            messagebox.showwarning(
                title="Nenhum produto selecionado",
                message="Clique em um produto da lista para remover.",
                parent=janela
            )

            return

        linha = selecao[0]
        produto = produtos_por_linha[linha]

        confirmar = messagebox.askyesno(
            title="Remover produto",
            message=f"Remover o produto \"{produto['nome']}\" da lista?",
            parent=janela
        )

        if not confirmar:
            return

        produtos.remove(produto)
        del produtos_por_linha[linha]
        tabela.delete(linha)

        if estado["editando"] == linha:
            sair_do_modo_edicao()


    # ==========================================
    # LISTA DE PRODUTOS ADICIONADOS
    # ==========================================

    label_lista = ctk.CTkLabel(
        janela,
        text="Produtos Adicionados  (duplo clique para editar)",
        font=("Arial", 16, "bold")
    )

    label_lista.pack(
        pady=(8, 4)
    )


    # ==========================================
    # BOTÕES (ficam presos embaixo da janela)
    # ==========================================
    # São colocados ANTES da tabela com side="bottom":
    # assim a tabela encolhe se faltar espaço, e os botões
    # nunca ficam cortados.

    frame_botoes = ctk.CTkFrame(
        janela,
        fg_color="transparent"
    )

    frame_botoes.pack(
        side="bottom",
        pady=(5, 15)
    )


    # ==========================================
    # TABELA
    # ==========================================

    frame_lista = ctk.CTkFrame(
        janela
    )

    frame_lista.pack(
        padx=40,
        pady=(0, 5),
        fill="both",
        expand=True
    )


    tabela = ttk.Treeview(
        frame_lista,
        columns=COLUNAS_COM_CODIGO,
        show="headings",
        selectmode="browse"
    )

    tabela.heading("sku", text="SKU")
    tabela.heading("nome", text="Produto")
    tabela.heading("marca", text="Marca")
    tabela.heading("codigo_barras", text="Cód. Barras")
    tabela.heading("unidade", text="Unidade")
    tabela.heading("preco", text="Preço")
    tabela.heading("estoque", text="Estoque")

    tabela.column("sku", width=120)
    tabela.column("nome", width=300)
    tabela.column("marca", width=140)
    tabela.column("codigo_barras", width=150)
    tabela.column("unidade", width=110)
    tabela.column("preco", width=100)
    tabela.column("estoque", width=100)


    # Barra de rolagem para listas grandes
    rolagem = ttk.Scrollbar(
        frame_lista,
        orient="vertical",
        command=tabela.yview
    )

    tabela.configure(yscrollcommand=rolagem.set)

    rolagem.pack(
        side="right",
        fill="y",
        pady=5
    )

    tabela.pack(
        fill="both",
        expand=True,
        padx=5,
        pady=5
    )


    # Duplo clique ou Enter na tabela abre a edição
    tabela.bind("<Double-1>", editar_selecionado)
    tabela.bind("<Return>", editar_selecionado)

    # Delete remove o item selecionado
    tabela.bind("<Delete>", lambda evento: remover_selecionado())


    # ==========================================
    # FUNÇÃO TERMINAR PLANILHA
    # ==========================================

    def terminar_planilha():

        if estado["editando"] is not None:

            messagebox.showwarning(
                title="Edição em andamento",
                message="Salve ou cancele a alteração antes de terminar a planilha.",
                parent=janela
            )

            return


        if not produtos:

            messagebox.showwarning(
                title="Nenhum produto",
                message="Adicione pelo menos um produto antes de terminar a planilha.",
                parent=janela
            )

            return


        try:

            caminho = gerar_planilha(produtos)

            messagebox.showinfo(
                title="Sucesso",
                message=(
                    f"Planilha gerada com sucesso!\n\n"
                    f"Produtos adicionados: {len(produtos)}\n\n"
                    f"Arquivo salvo em:\n{caminho}"
                ),
                parent=janela
            )

            # Abre a pasta já com o arquivo gerado selecionado (Windows)
            try:
                os.startfile(os.path.dirname(caminho))
            except Exception:
                pass

            janela.destroy()

        except Exception as erro:

            messagebox.showerror(
                title="Erro",
                message=f"Não foi possível gerar a planilha.\n\n{erro}",
                parent=janela
            )


    # ==========================================
    # BOTÕES
    # ==========================================

    botao_adicionar = ctk.CTkButton(
        frame_botoes,
        text="Adicionar Mais um Produto",
        width=240,
        height=45,
        command=adicionar_produto
    )

    botao_adicionar.grid(row=0, column=0, padx=8)


    botao_cancelar = ctk.CTkButton(
        frame_botoes,
        text="Cancelar Edição",
        width=160,
        height=45,
        fg_color="gray40",
        hover_color="gray30",
        command=sair_do_modo_edicao
    )

    botao_cancelar.grid(row=0, column=1, padx=8)
    botao_cancelar.grid_remove()   # só aparece durante a edição


    botao_editar = ctk.CTkButton(
        frame_botoes,
        text="Editar Selecionado",
        width=180,
        height=45,
        command=editar_selecionado
    )

    botao_editar.grid(row=0, column=2, padx=8)


    botao_remover = ctk.CTkButton(
        frame_botoes,
        text="Remover Selecionado",
        width=180,
        height=45,
        fg_color="#b33a3a",
        hover_color="#8f2d2d",
        command=remover_selecionado
    )

    botao_remover.grid(row=0, column=3, padx=8)


    botao_terminar = ctk.CTkButton(
        frame_botoes,
        text="Terminar Planilha",
        width=240,
        height=45,
        command=terminar_planilha
    )

    botao_terminar.grid(row=0, column=4, padx=8)


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