import tkinter as tk
from tkinter import ttk, messagebox, font
import sqlite3
import tkinter as tk


# Usado tkcalender para trabalhar com calendario e selecionar data
from tkcalendar import Calendar, DateEntry
from datetime import datetime
##### Implementação de sistema de relatório de produção gráfica para DCR
##### Trabalho de Estensão, 3º periodo.
##### Engloba discplinas Banco de Dados e Desenvolvimento Rápido de Aplicações em Pytho.
##### Professores Fabio Carneiro e Robsin Lorbieski
##### 09/2024
##### Por Carlos Henrique

class ProductManagementApp:
    def __init__(self, root):
        self.root = root
        self.root.title("RPI - Gráfica DCR (1.0)")
        
        # Desativa a função de maximizar
        root.resizable(False, False)

        # Centralizar janela principal
        window_width = root.winfo_reqwidth()
        window_height = root.winfo_reqheight()

        position_right = int(root.winfo_screenwidth()/4 - window_width/1)
        position_down = int(root.winfo_screenheight()/4 - window_height/2)

        root.geometry("+{}+{}".format(position_right, position_down))
        
        font_style = (10, "bold")
        font_style_titulo = (14, "bold", "italic")

        ##########################################################
        # Usada para filtrar Datas no tkcalender em search_data_tk
        self.label_resultado = tk.Label(root, text="") #Variavel para filtrar por data a busca por data

        # BARRA CINZA - menu #
        frame_janela1 = tk.Frame(root, width=150, height=30, bg="gray") # Titulo menu
        frame_janela1.grid(row=0, column=0, padx=5, pady=2, sticky="ew")
        frame_janela1 = tk.Frame(root, width=300, height=30, bg="gray") # Titulo menu
        frame_janela1.grid(row=0, column=1, padx=1, pady=2, sticky="w")

        # BARRA CINZA - Preenchimento
        frame_janela1 = tk.Frame(root, width=300, height=4, bg="#00008B") # Sub Titulo
        frame_janela1.grid(row=1, column=1, padx=1, pady=2, sticky="w")

        # BARRA CINZA - Area de Busca
        frame_janela1 = tk.Frame(root, width=360, height=4, bg="#00008B") # Sub Titulo
        frame_janela1.grid(row=1, column=2, padx=1, pady=2, sticky="w")

        # BARRA CINZA - Area de Busca
        #frame_janela1 = tk.Frame(root, width=240, height=4, bg="#00008B") # Sub Titulo
        #frame_janela1.grid(row=3, column=2, padx=1, pady=2, sticky="w")

        # BARRA CINZA - Somatorio das colunas
        frame_janela1 = tk.Frame(root, width=360, height=4, bg="#00008B") # Sub Somatorio das colunas
        frame_janela1.grid(row=10, column=2, padx=1, pady=2, sticky="w")

        # BARRA CINZA - Somatorio das colunas (coluna 3)
        frame_janela1 = tk.Frame(root, width=520, height=4, bg="#00008B") # Sub Somatorio das colunas
        frame_janela1.grid(row=10, column=3, padx=1, pady=2, sticky="w")

        # BARRA CINZA - Comandos
        frame_janela1 = tk.Frame(root, width=360, height=4, bg="#00008B") # Sub Comandos
        frame_janela1.grid(row=6, column=2, padx=1, pady=2, sticky="w")

        # BARRA CINZA - Codigo para horas improdutivas
        frame_janela1 = tk.Frame(root, width=520, height=4, bg="#00008B") 
        frame_janela1.grid(row=1, column=3, padx=1, pady=2, sticky="w")

        # BARRA CINZA - Tempo acerto
        frame_janela1 = tk.Frame(root, width=520, height=4, bg="#00008B") #Tempo acerto
        frame_janela1.grid(row=6, column=3, padx=1, pady=2, sticky="w")

        # barra para realçar alguns somatorios do treeview coluna 2
        frame_janela1 = tk.Frame(root, width=360, height=30, bg="gray") # resultado de calculo qpapel, hrum
        frame_janela1.grid(row=11, column=2, padx=10, pady=2, sticky="w")

        # barra para realçar alguns somatorios do treeview coluna 3
        frame_janela1 = tk.Frame(root, width=530, height=30, bg="gray") # resultado de calculo qpapel, hrum
        frame_janela1.grid(row=11, column=3, padx=5, pady=2, sticky="w")

        ############################
        # Conectar ao banco de dados
        self.conn = sqlite3.connect(r'D:\Meus Documetos\DCR_Estoque\SistemaDCR6_Correta\grafica_dcr2.db')
        self.cursor = self.conn.cursor()

                # Criar a tabela se não existir
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS RELAPROD (
                id INTEGER PRIMARY KEY,
                data TEXT NOT NULL,
                impressora TEXT NOT NULL,
                funcionario TEXT NOT NULL,                            
                cliente TEXT NOT NULL,
                n_of INTEGER NOT NULL,
                quant_cores INTEGER NOT NULL,
                medida TEXT NOT NULL,
                corte NUMERIC NOT NULL,
                quant_papel NUMERIC NOT NULL,
                metro_quadrado INTEGER NOT NULL,                            
                tipo_papel TEXT NOT NULL,
                horas_rum NUMERIC NOT NULL,
                ta_h NUMERIC NOT NULL,
                ta_i NUMERIC NOT NULL,
                hi_a NUMERIC NOT NULL,
                hi_b NUMERIC NOT NULL,
                hi_c NUMERIC NOT NULL,
                hi_d NUMERIC NOT NULL,
                hi_e NUMERIC NOT NULL,
                hi_f NUMERIC NOT NULL,
                hi_g NUMERIC NOT NULL,
                total_hp NUMERIC NOT NULL
            )
        ''')
        
        self.conn.commit()

        # Variáveis de controle
        self.selected_item = None

        # Interface
        self.create_widgets()

        # Carregar dados iniciais
        #self.fetch_data() #comentado para nao carregar o treeview na inicialização

        # Chamar a função clear_fields ao iniciar e limpa campos entry
        self.clear_fields() #comentado para nao carregar o treeview na inicialização

    def create_widgets(self):
        tk.Label(self.root, text="DCR Papers Rolls", font="font_style_titulo").grid(row=0, column=0, padx=(10, 1), pady=2, sticky="w")
        tk.Label(self.root, text="Relatório de Produção", font="font_style_titulo").grid(row=0, column=1, padx=(10, 1), pady=2, sticky="w")        
        tk.Label(self.root, text="Preencimento", font="font_style").grid(row=1, column=1, padx=10, pady=1, sticky="w")    #label para demarcar area de busca
        # Campos de entrada
        #################################################################
        ###### usado TKCALENDER para seleção de data no padrao DD/MM/AAAA
        tk.Label(self.root, text="Data").grid(row=2, column=0, padx=10, pady=2, sticky="e")
        self.entry_data = DateEntry(self.root, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd', state="readonly")
        self.entry_data.grid(row=2, column=1, padx=10, pady=2, columnspan=2, sticky="w")        

        ###############################################
        # Seleção em combox para qual Impressoras usada
        tk.Label(self.root, text="Impressora").grid(row=2, column=1, padx=(120, 1), pady=2, sticky="w")
        self.entry_tipo_impressora = ttk.Combobox(self.root, values=["MAQ 1", "MAQ 2"], width=12, state="readonly")
        self.entry_tipo_impressora.grid(row=2, column=1, padx=(185, 1), pady=2, sticky="w")

        # Seleção para funcinario que fez a impressão
        tk.Label(self.root, text="Funcionário").grid(row=3, column=0, padx=10, pady=2, sticky="e")
        self.entry_funcionario = ttk.Combobox(self.root, values=["JAILSON", "LEO", "LEO/JAILSON"], width=20, state="readonly")
        self.entry_funcionario.grid(row=3, column=1, padx=10, pady=2, sticky="w")

        tk.Label(self.root, text="Cliente").grid(row=4, column=0, padx=10, pady=2, sticky="e")
        self.entry_cliente = tk.Entry(self.root, width=45)
        self.entry_cliente.grid(row=4, column=1, padx=10, pady=2, sticky="w")

        tk.Label(self.root, text="N. OF").grid(row=5, column=0, padx=10, pady=2, sticky="e")
        self.entry_n_of = tk.Entry(self.root, width=15)
        self.entry_n_of.grid(row=5, column=1, padx=10, pady=2, sticky="w")

        # em quant cores, limitaado para ate 4 caracteres no label: irem validade_quant_cores / linha 165 aprox    
        tk.Label(self.root, text="Quant. Cores").grid(row=5, column=1, padx=(110, 1), pady=2, sticky="w")
        vcmd_papel = (self.root.register(self.validate_quant_cores), '%P')
        self.entry_quant_cores = tk.Entry(self.root, width=15, validate='key', validatecommand=vcmd_papel)
        self.entry_quant_cores.grid(row=5, column=1, padx=(190, 1), pady=2, sticky="w")

        tk.Label(self.root, text="Medida").grid(row=6, column=0, padx=10, pady=2, sticky="e")
        self.entry_medida = tk.Entry(self.root, width=10)
        self.entry_medida.grid(row=6, column=1, padx=10, pady=2, sticky="w")

        tk.Label(self.root, text="Corte (mm)").grid(row=6, column=1, padx=(80, 1), pady=2, sticky="w")
        self.entry_corte = tk.Entry(self.root, width=8)
        self.entry_corte.grid(row=6, column=1, padx=(150, 1), pady=2, sticky="w")

        tk.Label(self.root, text="Q. Papel (m)").grid(row=7, column=0, padx=10, pady=2, sticky="e")
        self.entry_quant_papel = tk.Entry(self.root, width=10)
        self.entry_quant_papel.grid(row=7, column=1, padx=10, pady=2, sticky="w")

        ######################################################################
        # LABEL E CX DE ENTRADA PARA EXIBIR CALCULO DOS CAMPOS METRO, QUADRADO
        tk.Label(self.root, text="m²").grid(row=6, column=1, padx=(210, 1), pady=2, sticky="w")
        self.entry_metro_quadrado = tk.Entry(self.root, width=8, state='normal')
        self.entry_metro_quadrado.grid(row=6, column=1, padx=(230, 1), pady=2, sticky="w")
        
        ####################################################################
        #label para demarcar somatorio        
        tk.Label(self.root, text="Somatório de Colunas", font="font_style").grid(row=10, column=2, padx=10, pady=1, sticky="w")  
        # Label para exibir o SOMATORIO. Posição / Aparece acima do TREEVIEW        
        self.somatorio_quant_papel_label = tk.Label(self.root, text="Q.Papel:")
        self.somatorio_quant_papel_label.grid(row=11, column=2, columnspan=2, padx=(20,1), pady=2, sticky="w")        

        self.somatorio_horas_rum_label = tk.Label(self.root, text="H.Rum:")
        self.somatorio_horas_rum_label.grid(row=11, column=2, columnspan=2, padx=(120,1), pady=2, sticky="w")        

        self.somatorio_ta_h_label = tk.Label(self.root, text="TA (H):")
        self.somatorio_ta_h_label.grid(row=11, column=2, columnspan=2, padx=(190, 1), pady=2, sticky="w")        

        self.somatorio_ta_i_label = tk.Label(self.root, text="TA (I):")
        self.somatorio_ta_i_label.grid(row=11, column=2, columnspan=2, padx=(260, 1), pady=2, sticky="w")        

        self.somatorio_hi_a_label = tk.Label(self.root, text="A:")
        self.somatorio_hi_a_label.grid(row=11, column=2, columnspan=2, padx=(330, 1), pady=2, sticky="w")        

        self.somatorio_hi_b_label = tk.Label(self.root, text="B:")
        self.somatorio_hi_b_label.grid(row=11, column=3, columnspan=2, padx=10, pady=2, sticky="w")        

        self.somatorio_hi_c_label = tk.Label(self.root, text="C:")
        self.somatorio_hi_c_label.grid(row=11, column=3, columnspan=2, padx=(60, 1), pady=2, sticky="w")        

        self.somatorio_hi_d_label = tk.Label(self.root, text="D:")
        self.somatorio_hi_d_label.grid(row=11, column=3, columnspan=2, padx=(110, 1), pady=2, sticky="w")        

        self.somatorio_hi_e_label = tk.Label(self.root, text="E:")
        self.somatorio_hi_e_label.grid(row=11, column=3, columnspan=2, padx=(160, 1), pady=2, sticky="w")        

        self.somatorio_hi_f_label = tk.Label(self.root, text="G:")
        self.somatorio_hi_f_label.grid(row=11, column=3, columnspan=2, padx=(210, 1), pady=2, sticky="w")        

        self.somatorio_hi_g_label = tk.Label(self.root, text="G:")
        self.somatorio_hi_g_label.grid(row=11, column=3, columnspan=2, padx=(260, 1), pady=2, sticky="w")        

        self.somatorio_total_hp_label = tk.Label(self.root, text="Total Horas Paga")
        self.somatorio_total_hp_label.grid(row=11, column=3, columnspan=10, padx=(400, 1), pady=2, sticky="w")        


        ####################################################
        # ASSOCIANDO EVENTOS DE PARA CALCULAR METRO QUADRADO
        self.entry_corte.bind("<KeyRelease>", self.calculate_metro_quadrado)
        self.entry_quant_papel.bind("<KeyRelease>", self.calculate_metro_quadrado)

        ######################################
        # Escolha de Tipo de papel com Combox
        tk.Label(self.root, text="Tipo Papel").grid(row=7, column=1, padx=(85, 1), pady=2, sticky="w")
        self.entry_tipo_papel = ttk.Combobox(self.root, values=["BOPP BRANCO", "BOPP METALIZADO", "COUCHE","FILME PICOLE 26g", "FILME PICOLE 40g", "TERMICO"], width=18, state="readonly")
        self.entry_tipo_papel.grid(row=7, column=1, padx=(150, 1), pady=2, sticky="w")
        
        # validate_pontovirgula, Valida Entrada somente com ponto / horas rum + outros campos numericos decimais ....
        tk.Label(self.root, text="Horas Rum").grid(row=8, column=0, padx=10, pady=2, sticky="e")
        vcmd = (root.register(self.validate_pontovirgula), '%P')
        self.entry_horas_rum = tk.Entry(self.root, width=10, validate="key", validatecommand=vcmd)
        self.entry_horas_rum.grid(row=8, column=1, padx=10, pady=2, sticky="w")

        ##########################################################
        # ASSOCIANDO EVENTOS PARA CALCULAR HORAS PAGAS total_hp
        self.entry_horas_rum.bind("<KeyRelease>", self.calculate_total_hp)

        ###########################################
        # label, entrada de TEMPO ACERTO (h e i)
        tk.Label(self.root, text="Tempo de Acerto", font="font_style").grid(row=9, column=0, padx=10, pady=2, sticky="e")

        tk.Label(self.root, text="(H)").grid(row=9, column=1, padx=10, pady=2, sticky="w")
        vcmd = (root.register(self.validate_pontovirgula), '%P')        
        self.entry_ta_h = tk.Entry(self.root, width=5, validate="key", validatecommand=vcmd)
        self.entry_ta_h.grid(row=9, column=1, padx=(35, 1), pady=2, sticky="w")

        # ASSOCIANDO EVENTOS DE PARA CALCULAR HORAS PAGAS total_hp
        self.entry_ta_h.bind("<KeyRelease>", self.calculate_total_hp)

        tk.Label(self.root, text="(I)").grid(row=9, column=1, padx=(70, 1), pady=2, sticky="w")
        vcmd = (root.register(self.validate_pontovirgula), '%P')                
        self.entry_ta_i = tk.Entry(self.root, width=5, validate="key", validatecommand=vcmd)
        self.entry_ta_i.grid(row=9, column=1, padx=(90, 1), pady=2, sticky="w")

        # ASSOCIANDO EVENTOS DE PARA CALCULAR HORAS PAGAS total_hp
        self.entry_ta_i.bind("<KeyRelease>", self.calculate_total_hp)

        ###############################################
        # label, entrada de HORAS IMPRODUTIVAS (a ate g)
        tk.Label(self.root, text="Horas Improdutivas", font="font_style").grid(row=10, column=0, padx=10, pady=2, sticky="e")
        
        tk.Label(self.root, text="A").grid(row=10, column=1, padx=10, pady=2, sticky="w")
        vcmd = (root.register(self.validate_pontovirgula), '%P')                        
        self.entry_hi_a = tk.Entry(self.root, width=5, validate="key", validatecommand=vcmd)
        self.entry_hi_a.grid(row=11, column=1, padx=10, pady=2, sticky="w")
        
        # ASSOCIANDO EVENTOS DE PARA CALCULAR HORAS PAGAS total_hp
        self.entry_hi_a.bind("<KeyRelease>", self.calculate_total_hp)

        tk.Label(self.root, text="B").grid(row=10, column=1, padx=(50, 1), pady=2, sticky="w")
        vcmd = (root.register(self.validate_pontovirgula), '%P')                                
        self.entry_hi_b = tk.Entry(self.root, width=5, validate="key", validatecommand=vcmd)
        self.entry_hi_b.grid(row=11, column=1, padx=(50, 1), pady=2, sticky="w")

        # ASSOCIANDO EVENTOS DE PARA CALCULAR HORAS PAGAS total_hp
        self.entry_hi_b.bind("<KeyRelease>", self.calculate_total_hp)

        tk.Label(self.root, text="C").grid(row=10, column=1, padx=(90, 1), pady=2, sticky="w")
        vcmd = (root.register(self.validate_pontovirgula), '%P')        
        self.entry_hi_c = tk.Entry(self.root, width=5, validate="key", validatecommand=vcmd)
        self.entry_hi_c.grid(row=11, column=1, padx=(90, 1), pady=2, sticky="w")

        # ASSOCIANDO EVENTOS DE PARA CALCULAR HORAS PAGAS total_hp
        self.entry_hi_c.bind("<KeyRelease>", self.calculate_total_hp)

        tk.Label(self.root, text="D").grid(row=10, column=1, padx=(130, 1), pady=2, sticky="w")
        vcmd = (root.register(self.validate_pontovirgula), '%P')                
        self.entry_hi_d = tk.Entry(self.root, width=5, validate="key", validatecommand=vcmd)
        self.entry_hi_d.grid(row=11, column=1, padx=(130, 1), pady=2, sticky="w")

        # ASSOCIANDO EVENTOS DE PARA CALCULAR HORAS PAGAS total_hp
        self.entry_hi_d.bind("<KeyRelease>", self.calculate_total_hp)

        tk.Label(self.root, text="E").grid(row=10, column=1, padx=(170, 1), pady=2, sticky="w")
        vcmd = (root.register(self.validate_pontovirgula), '%P')                        
        self.entry_hi_e = tk.Entry(self.root, width=5, validate="key", validatecommand=vcmd)
        self.entry_hi_e.grid(row=11, column=1, padx=(170, 1), pady=2, sticky="w")

        # ASSOCIANDO EVENTOS DE PARA CALCULAR HORAS PAGAS total_hp
        self.entry_hi_e.bind("<KeyRelease>", self.calculate_total_hp)

        tk.Label(self.root, text="F").grid(row=10, column=1, padx=(210, 1), pady=2, sticky="w")
        vcmd = (root.register(self.validate_pontovirgula), '%P')                                
        self.entry_hi_f = tk.Entry(self.root, width=5, validate="key", validatecommand=vcmd)
        self.entry_hi_f.grid(row=11, column=1, padx=(210, 1), pady=2, sticky="w")

        # ASSOCIANDO EVENTOS DE PARA CALCULAR HORAS PAGAS total_hp
        self.entry_hi_f.bind("<KeyRelease>", self.calculate_total_hp)

        tk.Label(self.root, text="G").grid(row=10, column=1, padx=(250, 1), pady=2, sticky="w")
        vcmd = (root.register(self.validate_pontovirgula), '%P')                                
        self.entry_hi_g = tk.Entry(self.root, width=5, validate="key", validatecommand=vcmd)
        self.entry_hi_g.grid(row=11, column=1, padx=(250, 1), pady=2, sticky="w")

        # ASSOCIANDO EVENTOS DE PARA CALCULAR HORAS PAGAS total_hp
        self.entry_hi_g.bind("<KeyRelease>", self.calculate_total_hp)

        # Soma de tempos TOTAL HORAS PAGAS ###################
        tk.Label(self.root, text="Total HS PG").grid(row=8, column=1, padx=(78, 1), pady=2, sticky="w")
        self.entry_total_hp = tk.Entry(self.root, width=21, state='readonly')
        self.entry_total_hp.grid(row=8, column=1, padx=(150, 1), pady=2, sticky="w")

        # ASSOCIANDO EVENTOS DE PARA CALCULAR HORAS PAGAS total_hp
        self.entry_total_hp.bind("<KeyRelease>", self.calculate_total_hp)

        self.bind_events()

        #######################
        # BOTOES DE COMANDO
        tk.Label(self.root, text="Comandos", font="font_style").grid(row=6, column=2, padx=10, pady=1, sticky="w")    #label para demarcar Botôes

        self.btn_cadastrar = tk.Button(self.root, width=15, text="Cadastrar", command=self.insert_data)
        self.btn_cadastrar.grid(row=7, column=2, padx=10, pady=2, sticky="w")

        self.btn_atualizar = tk.Button(self.root, width=15, text="Alterar", command=self.update_data, state=tk.DISABLED) # Regra do Atualizar
        self.btn_atualizar.grid(row=7, column=2, padx=(130, 1), pady=2, sticky="w")

        self.btn_limpar = tk.Button(self.root, width=15, text="Limpa/Esc", command=self.clear_fields_2)
        self.btn_limpar.grid(row=8, column=2, padx=10, pady=2, sticky="w")

        self.btn_apagar = tk.Button(self.root, width=15, text="Apagar", command=self.delete_data, state=tk.DISABLED)
        self.btn_apagar.grid(row=8, column=2, padx=(130, 1), pady=2, sticky="w")

        # Desabilitado botão calculo
        self.btn_calcular = tk.Button(self.root,  width=15, text="Amostrar Tudo", command=self.amostrar_tudo)
        self.btn_calcular.grid(row=8, column=2, padx=(250, 1), pady=2, sticky="w")

        #####################################################################
        ### seleção de data para filtragem por periodo com TKCALENDER - Busca - variavel usada label_resultado
        tk.Label(self.root, text="Busca combinada - Texto e/ou Período", font="font_style").grid(row=1, column=2, padx=10, pady=1, sticky="w")    #label para demarcar area de busca

        #tk.Label(self.root, text="Buscar por período", font="font_style").grid(row=3, column=2, padx=10, pady=1, sticky="w")    
        tk.Label(self.root, text="Data Inicial").grid(row=3, column=2, padx=10, pady=1, sticky="w")
        self.entry_data_inicial = DateEntry(self.root, width=15, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd', state="normal")
        self.entry_data_inicial.grid(row=4, column=2, padx=10, pady=1, columnspan=2, sticky="w")

        tk.Label(self.root, text="Data Final").grid(row=3, column=2, padx=(130, 10), pady=2, sticky="w")
        self.entry_data_final = DateEntry(self.root, width=15, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd', state="normal")
        self.entry_data_final.grid(row=4, column=2, padx=(130, 10), pady=2, columnspan=2, sticky="w")

        #################
        # Campo de busca
        tk.Label(self.root, text="Buscar texto").grid(row=2, column=2, padx=(10, 1), pady=5, sticky="w") #label
        self.entry_search = tk.Entry(self.root, width=45) #Caixa de texto de busca
        self.entry_search.grid(row=2, column=2, padx=(82, 1), pady=10, sticky="w") #caixa de texto busca

        self.btn_buscar = tk.Button(self.root, text="Buscar Geral", command=self.search_data, width=15) #botao busca por texto
        self.btn_buscar.grid(row=4, column=2, padx=(250, 1), pady=10, sticky="w")

        #####################################################################
        # SUMARIO # explicação dos codigos - Códigos para horas improdutivas
        tk.Label(self.root, text="Codigos para horas improdutivas:", font="font_style").grid(row=1, column=3, padx=10, pady=5, sticky="w") #Titulo
        tk.Label(self.root, text="(A) - Falta/café/férias/outras faltas").grid(row=2, column=3, padx=1, pady=5, sticky="w") #label        
        tk.Label(self.root, text="(B) - Aguardando matérial/Serviço").grid(row=3, column=3, padx=1, pady=5, sticky="w") #label                
        tk.Label(self.root, text="(C) - Falta de energia").grid(row=4, column=3, padx=1, pady=5, sticky="w") #label                        
        tk.Label(self.root, text="(D) - Manutenção máquina").grid(row=5, column=3, padx=1, pady=5, sticky="w") #label
        tk.Label(self.root, text="(E) - A serviço do almoxarifado").grid(row=2, column=3, padx=(250, 1), pady=5, sticky="w") #label                                
        tk.Label(self.root, text="(F) - Treinamento").grid(row=3, column=3, padx=(250, 1), pady=5, sticky="w") #label                                
        tk.Label(self.root, text="(G) - Outros (Especificar)").grid(row=4, column=3, padx=(250, 1), pady=5, sticky="w") #label 
        tk.Label(self.root, text="Tempo de Acerto", font="font_style").grid(row=6, column=3, padx=10, pady=5, sticky="w") # titulo
        tk.Label(self.root, text="TA (H) - Lavagem cilidro").grid(row=7, column=3, padx=1, pady=5, sticky="w") #label 
        tk.Label(self.root, text="TA (I) - Montagem cilidro e ajuste papel").grid(row=8, column=3, padx=1, pady=1, sticky="w") #label 
        tk.Label(self.root, text="Horas Rum - Tempo de Impressão").grid(row=9, column=3, padx=1, pady=1, sticky="w") #label 
        tk.Label(self.root, text="T.HP - Total horas paga").grid(row=7, column=3, padx=(250, 1), pady=1, sticky="w") #label         

        ###############################
        # Treeview para exibir os dados
        self.tree = ttk.Treeview(root, columns=("ID", "Data", "Impressora", "Funcionario", "Cliente", "N. OF", "Quant. Cores", "Medida", "Corte", "Quant. Papel", "metro_quadrado", "Tipo Papel", "Horas Rum", "ta_h", "ta_i", "hi_a", "hi_b", "hi_c", "hi_d", "hi_e", "hi_f", "hi_g", "total_hp"), show="headings", height=12)
        self.tree.heading("ID", text="ID")
        self.tree.heading("Data", text="Data")
        self.tree.heading("Impressora", text="Impressora")
        self.tree.heading("Funcionario", text="Funcionario")        
        self.tree.heading("Cliente", text="Cliente")
        self.tree.heading("N. OF", text="N. OF")
        self.tree.heading("Quant. Cores", text="Cores")
        self.tree.heading("Medida", text="Medida")
        self.tree.heading("Corte", text="Corte")
        self.tree.heading("Quant. Papel", text="Q. Papel")
        self.tree.heading("metro_quadrado", text="m²")        
        self.tree.heading("Tipo Papel", text="Tipo Papel")
        self.tree.heading("Horas Rum", text="H. RUM")
        self.tree.heading("ta_h", text="TA (H)")  
        self.tree.heading("ta_i", text="TA (I)") 
        self.tree.heading("hi_a", text="A") 
        self.tree.heading("hi_b", text="B") 
        self.tree.heading("hi_c", text="C")         
        self.tree.heading("hi_d", text="D")        
        self.tree.heading("hi_e", text="E")                
        self.tree.heading("hi_f", text="F") 
        self.tree.heading("hi_g", text="G")   
        self.tree.heading("total_hp", text="T.HP ")           
        self.tree.column("ID", width=10)
        self.tree.column("Data", width=50)
        self.tree.column("Impressora", width=45)        
        self.tree.column("Funcionario", width=60)
        self.tree.column("Cliente", width=140)
        self.tree.column("N. OF", width=25)
        self.tree.column("Quant. Cores", width=25)
        self.tree.column("Medida", width=40)
        self.tree.column("Corte", width=25)
        self.tree.column("Quant. Papel", width=40)
        self.tree.column("metro_quadrado", width=30)        
        self.tree.column("Tipo Papel", width=90)
        self.tree.column("Horas Rum", width=30)
        self.tree.column("ta_h", width=18)
        self.tree.column("ta_i", width=18)
        self.tree.column("hi_a", width=10) 
        self.tree.column("hi_b", width=10)         
        self.tree.column("hi_c", width=10)         
        self.tree.column("hi_d", width=10)        
        self.tree.column("hi_e", width=10)                
        self.tree.column("hi_f", width=10)                
        self.tree.column("hi_g", width=10)
        self.tree.column("total_hp", width=20)

        self.tree.grid(row=17, column=0, columnspan=10, padx=10, pady=10, ipadx=80)

        #alinhamento no treeview
        self.tree.column("ID", anchor="center")
        self.tree.column("Cliente", anchor="w")

        # Configurar evento de clique duplo no Treeview
        self.tree.bind("<Double-1>", self.on_double_click)

        # Barra de rolagem vertical
        scrollbar = ttk.Scrollbar(self.root, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar.grid(row=17, column=10, padx=(1, 1), sticky='ns')
        self.tree.configure(yscrollcommand=scrollbar.set)

        ################################################################################
        # Aqui começa teste de campos para ponto no decimal no lugar de virgula (. / ,)
        ################################################################################        
    def validate_pontovirgula(self, new_value): # Permite valores vazios ou números que contenham apenas um ponto decimal
        if new_value == "" or new_value.replace('.', '', 1).isdigit():
           return True
        else:
            messagebox.showwarning("Valor Inválido", "Por favor, insira um número válido com ponto decimal (.) e não com vírgula (,).")
        return False

    # Somatorios de colunas para resultados em label acima do TreeView
    def update_somatorio(self): #aqui os somatorios
        somatorio_horas_rum = 0
        somatorio_quant_papel = 0
        somatorio_ta_h = 0
        somatorio_ta_i = 0  
        somatorio_hi_a = 0   
        somatorio_hi_b = 0 
        somatorio_hi_c = 0         
        somatorio_hi_d = 0
        somatorio_hi_e = 0        
        somatorio_hi_f = 0        
        somatorio_hi_g = 0
        somatorio_total_hp = 0        
        for child in self.tree.get_children():
            item = self.tree.item(child)['values']
            horas_rum = item[12]  # horas_rum (Verificar sempre numero da coluna, caso tenha alteração)
            quant_papel = item[9]  # quant_papel (Verificar sempre numero da coluna, caso tenha alteração)
            ta_h = item[13]
            ta_i = item[14]
            hi_a = item[15]
            hi_b = item[16]
            hi_c = item[17]            
            hi_d = item[18]
            hi_e = item[19]            
            hi_f = item[20] 
            hi_g = item[21] 
            total_hp = item[22]             
            somatorio_horas_rum += float(horas_rum)
            somatorio_quant_papel += float(quant_papel)
            somatorio_ta_h += float(ta_h) 
            somatorio_ta_i += float(ta_i) 
            somatorio_hi_a += float(hi_a)             
            somatorio_hi_b += float(hi_b) 
            somatorio_hi_c += float(hi_c)
            somatorio_hi_d += float(hi_d)
            somatorio_hi_e += float(hi_e)
            somatorio_hi_f += float(hi_f)  
            somatorio_hi_g += float(hi_g)              
            somatorio_total_hp += float(total_hp) 
            
            # Atualiza o texto do Label do somatório
            self.somatorio_horas_rum_label.config(text=f"H.Rum: {somatorio_horas_rum:.1f}") #  HR: é o label na tela, Exibe numero decimal
            self.somatorio_quant_papel_label.config(text=f"Q.Papel: {int(somatorio_quant_papel)}") #  QP: é o label na tela, exibe numero inteiro
            self.somatorio_ta_h_label.config(text=f"TA (H): {somatorio_ta_h:.1f}") #  Exibe somatorio como Decimal, 2 digitos apos o ponto
            self.somatorio_ta_i_label.config(text=f"TA (I): {somatorio_ta_i:.1f}") 
            self.somatorio_hi_a_label.config(text=f"A: {somatorio_hi_a:.1f}")             
            self.somatorio_hi_b_label.config(text=f"B: {somatorio_hi_b:.1f}") 
            self.somatorio_hi_c_label.config(text=f"C: {somatorio_hi_c:.1f}")
            self.somatorio_hi_d_label.config(text=f"D: {somatorio_hi_d:.1f}")            
            self.somatorio_hi_e_label.config(text=f"E: {somatorio_hi_e:.1f}") 
            self.somatorio_hi_f_label.config(text=f"F: {somatorio_hi_f:.1f}")            
            self.somatorio_hi_g_label.config(text=f"G: {somatorio_hi_g:.1f}") 
            self.somatorio_total_hp_label.config(text=f"Total Horas Paga: {somatorio_total_hp:.1f}")             

        # Limita o campo quant_papel (quantidade cores) a 4 caracteres
    def validate_quant_cores(self, P):
        return len(P) <= 2

        # Atualiza treeview com dados do banco atraves do etch_dataf
    def fetch_data(self):
        self.cursor.execute('SELECT * FROM RELAPROD')
        rows = self.cursor.fetchall()
        self.tree.delete(*self.tree.get_children())
        for row in rows:
            self.tree.insert("", tk.END, values=row)

    # calcula O metro quadrado a partir dos capos CORTE e QUANT_PAPEL
    def calculate_metro_quadrado(self, event):
        try:
            corte = float(self.entry_corte.get().replace(",", "."))
            quant_papel = float(self.entry_quant_papel.get())
            metro_quadrado = (corte / 1000) * quant_papel
            self.entry_metro_quadrado.config(state='normal')
            self.entry_metro_quadrado.delete(0, tk.END)
            self.entry_metro_quadrado.insert(0, f"{metro_quadrado:.2f}")
            self.entry_metro_quadrado.config(state='disable')
        except ValueError:
            self.entry_metro_quadrado.config(state='normal')
            self.entry_metro_quadrado.delete(0, tk.END)
            self.entry_metro_quadrado.insert(0, "0")
            self.entry_metro_quadrado.config(state='disable')

    # TOTAL HORAS PAGAS total_hp
    def calculate_total_hp(self, event):
        try:
            horas_rum = float(self.entry_horas_rum.get().replace(",", "."))
            ta_h = float(self.entry_ta_h.get().replace(",", "."))
            ta_i = float(self.entry_ta_i.get().replace(",", "."))            
            hi_a = float(self.entry_hi_a.get().replace(",", ".")) 
            hi_b = float(self.entry_hi_b.get().replace(",", "."))             
            hi_c = float(self.entry_hi_c.get().replace(",", "."))             
            hi_d = float(self.entry_hi_d.get().replace(",", "."))             
            hi_e = float(self.entry_hi_e.get().replace(",", "."))             
            hi_f = float(self.entry_hi_f.get().replace(",", "."))             
            hi_g = float(self.entry_hi_g.get().replace(",", "."))             
            total_hp = horas_rum + ta_h + ta_i + hi_a + hi_b + hi_c + hi_d + hi_e + hi_f + hi_g
            self.entry_total_hp.config(state='normal')
            self.entry_total_hp.delete(0, tk.END)
            self.entry_total_hp.insert(0, f"{total_hp:.2f}")
            self.entry_total_hp.config(state='disable')
        except ValueError:
            self.entry_total_hp.config(state='normal')
            self.entry_total_hp.delete(0, tk.END)
            self.entry_total_hp.insert(0, "0")
            self.entry_total_hp.config(state='disable')

    def bind_events(self):
        self.entry_horas_rum.bind("<KeyRelease>", self.calculate_total_hp)
        self.entry_ta_h.bind("<KeyRelease>", self.calculate_total_hp)
        self.entry_ta_i.bind("<KeyRelease>", self.calculate_total_hp)
        self.entry_hi_a.bind("<KeyRelease>", self.calculate_total_hp)
        self.entry_hi_b.bind("<KeyRelease>", self.calculate_total_hp)
        self.entry_hi_c.bind("<KeyRelease>", self.calculate_total_hp)    
        self.entry_hi_d.bind("<KeyRelease>", self.calculate_total_hp)    
        self.entry_hi_e.bind("<KeyRelease>", self.calculate_total_hp)    
        self.entry_hi_f.bind("<KeyRelease>", self.calculate_total_hp)
        self.entry_total_hp.bind("<KeyRelease>", self.calculate_total_hp)
    
   # insere dados no banco
    def insert_data(self):
        data = self.entry_data.get().strip()
        impressora = self.entry_tipo_impressora.get().strip().upper()
        funcionario = self.entry_funcionario.get().strip().upper()        
        cliente = self.entry_cliente.get().strip().upper()
        n_of = self.entry_n_of.get().strip().upper()
        quant_cores = self.entry_quant_cores.get().strip()
        medida = self.entry_medida.get().strip().upper()
        corte = self.entry_corte.get().strip().upper()
        quant_papel = self.entry_quant_papel.get().strip()
        metro_quadrado = self.entry_metro_quadrado.get().strip()
        tipo_papel = self.entry_tipo_papel.get().strip().upper()
        horas_rum = self.entry_horas_rum.get().strip()  
        ta_h = self.entry_ta_h.get().strip()
        ta_i = self.entry_ta_i.get().strip()        
        hi_a = self.entry_hi_a.get().strip()        
        hi_b = self.entry_hi_b.get().strip()        
        hi_c = self.entry_hi_c.get().strip()
        hi_d = self.entry_hi_d.get().strip()
        hi_e = self.entry_hi_e.get().strip()
        hi_f = self.entry_hi_f.get().strip()        
        hi_g = self.entry_hi_g.get().strip()                
        total_hp = self.entry_total_hp.get().strip()  

        if (data and impressora and funcionario and cliente and n_of.isdigit() and quant_cores.isdigit() and medida and corte.isdigit() and quant_papel.replace('.', '', 1).isdigit() and metro_quadrado and tipo_papel and horas_rum.replace(',', '', 1) and ta_h.replace(',', '', 1) and ta_i.replace(',', '', 1) and hi_a.replace(',', '', 1) and hi_b.replace(',', '', 1) and hi_c.replace(',', '', 1) and hi_d.replace(',', '', 1) and hi_e.replace(',', '', 1) and hi_f.replace(',', '', 1) and hi_g.replace(',', '', 1) and total_hp):
            self.cursor.execute('''
                INSERT INTO RELAPROD (data, impressora, funcionario, cliente, n_of, quant_cores, medida, corte, quant_papel, metro_quadrado, tipo_papel, horas_rum, ta_h, ta_i, hi_a, hi_b, hi_c, hi_d, hi_e, hi_f, hi_g, total_hp)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (data, impressora, funcionario, cliente, n_of, quant_cores, medida, corte, quant_papel, metro_quadrado, tipo_papel, horas_rum, ta_h, ta_i, hi_a, hi_b, hi_c, hi_d, hi_e, hi_f, hi_g, total_hp))
            self.conn.commit()
                
            self.fetch_data()
            self.clear_fields()
            # Atualizar o somatório após inserir os dados
            self.update_somatorio()
        else:
            messagebox.showerror("Erro", "Por favor, preencha todos os campos corretamente.\nAtenção aos Campos que aceitam apenas Numeros.")


    # A função de atualização , atualiza o banco.
    def update_data(self):
        if self.selected_item:
            data = self.entry_data.get().strip()
            impressora = self.entry_tipo_impressora.get().strip().upper()
            funcionario = self.entry_funcionario.get().strip().upper()
            cliente = self.entry_cliente.get().strip().upper()
            n_of = self.entry_n_of.get().strip().upper()
            quant_cores = self.entry_quant_cores.get().strip()
            medida = self.entry_medida.get().strip().upper()
            corte = self.entry_corte.get().strip().upper()
            quant_papel = self.entry_quant_papel.get().strip()
            metro_quadrado = self.entry_metro_quadrado.get().strip()
            tipo_papel = self.entry_tipo_papel.get().strip().upper()
            horas_rum = self.entry_horas_rum.get().strip() 
            ta_h = self.entry_ta_h.get().strip() 
            ta_i = self.entry_ta_i.get().strip() 
            hi_a = self.entry_hi_a.get().strip()             
            hi_b = self.entry_hi_b.get().strip()
            hi_c = self.entry_hi_c.get().strip()
            hi_d = self.entry_hi_d.get().strip()
            hi_e = self.entry_hi_e.get().strip()            
            hi_f = self.entry_hi_f.get().strip()            
            hi_g = self.entry_hi_g.get().strip()            
            total_hp = self.entry_total_hp.get().strip()                        

            if (data and impressora and funcionario and cliente and n_of.isdigit() and quant_cores.isdigit() and medida and corte.isdigit() and quant_papel.replace('.', '', 1).isdigit() and metro_quadrado and tipo_papel and horas_rum.replace(',', '', 1) and ta_h.replace(',', '', 1) and ta_i.replace(',', '', 1) and hi_a.replace(',', '', 1) and hi_b.replace(',', '', 1) and hi_c.replace(',', '', 1) and hi_d.replace(',', '', 1) and hi_e.replace(',', '', 1) and hi_f.replace(',', '', 1) and hi_g.replace(',', '', 1) and total_hp):
                self.cursor.execute('''
                    UPDATE RELAPROD
                    SET data=?, impressora=?, funcionario=?, cliente=?, n_of=?, quant_cores=?, medida=?, corte=?, quant_papel=?, metro_quadrado=?, tipo_papel=?, horas_rum=?, ta_h=?, ta_i=?, hi_a=?, hi_b=?, hi_c=?, hi_d=?, hi_e=?, hi_f=?, hi_g=?, total_hp=?
                    WHERE id=?
                ''', (data, impressora, funcionario, cliente, n_of, quant_cores, medida, corte, quant_papel, metro_quadrado, tipo_papel, horas_rum, ta_h, ta_i, hi_a, hi_b, hi_c, hi_d, hi_e, hi_f, hi_g, total_hp, self.selected_item['values'][0]))
                self.conn.commit()
                self.fetch_data()
                self.clear_fields()
                self.btn_cadastrar.config(state=tk.NORMAL)
                self.btn_atualizar.config(state=tk.DISABLED)
                self.btn_apagar.config(state=tk.DISABLED)
                self.btn_calcular.config(state=tk.NORMAL)            
                self.selected_item = None
                # Atualizar o somatório após atualizar os dados
                self.update_somatorio()
            else:
                messagebox.showerror("Erro", "Por favor, preencha todos os campos corretamente.\nAtenção aos Campos que aceitam apenas Numeros.")
        else:
            messagebox.showerror("Erro", "Selecione uma linha para atualizar.")



    def delete_data(self):
        if self.selected_item:
            confirmar = messagebox.askyesno("Confirmar", "Deseja realmente apagar este cadastro?")
            if confirmar:
                self.cursor.execute('DELETE FROM RELAPROD WHERE id=?', (self.selected_item['values'][0],))
                self.conn.commit()
                self.fetch_data()
                self.clear_fields()
                self.btn_cadastrar.config(state=tk.NORMAL)
                self.btn_atualizar.config(state=tk.DISABLED)
                self.btn_apagar.config(state=tk.DISABLED)
                self.selected_item = None
                # Atualizar o somatório após deletar os dados
                self.update_somatorio()
        else:
            messagebox.showerror("Erro", "Selecione um item para apagar.")

    # Limpa texto dos campos

    def clear_fields(self): #limpeza padrao acionada por outros def
        self.entry_data.delete(0, tk.END)
        self.entry_tipo_impressora.config(state='normal')        
        self.entry_tipo_impressora.delete(0, tk.END)
        self.entry_tipo_impressora.config(state='readonly')        
        self.entry_funcionario.config(state='normal')
        self.entry_funcionario.delete(0, tk.END)
        self.entry_funcionario.config(state='readonly')        
        self.entry_cliente.delete(0, tk.END)
        self.entry_n_of.delete(0, tk.END)
        self.entry_quant_cores.delete(0, tk.END)
        self.entry_medida.delete(0, tk.END)
        self.entry_corte.delete(0, tk.END)
        self.entry_quant_papel.delete(0, tk.END)
        self.entry_metro_quadrado.config(state='normal')
        self.entry_metro_quadrado.delete(0, tk.END)
        self.entry_metro_quadrado.config(state='disable') 
        self.entry_tipo_papel.config(state='normal')                  
        self.entry_tipo_papel.delete(0, tk.END)
        self.entry_tipo_papel.config(state='readonly')
        self.entry_horas_rum.delete(0, tk.END)
        self.entry_ta_h.delete(0, tk.END)        
        self.entry_ta_i.delete(0, tk.END) 
        self.entry_hi_a.delete(0, tk.END)         
        self.entry_hi_b.delete(0, tk.END)
        self.entry_hi_c.delete(0, tk.END)
        self.entry_hi_d.delete(0, tk.END) 
        self.entry_hi_e.delete(0, tk.END)         
        self.entry_hi_f.delete(0, tk.END)         
        self.entry_hi_g.delete(0, tk.END)
        self.entry_total_hp.config(state='normal')        
        self.entry_total_hp.delete(0, tk.END)
        self.entry_total_hp.config(state='disable')        
        self.entry_data_inicial.delete(0, tk.END)
        self.entry_data_final.delete(0, tk.END)

        self.btn_cadastrar.config(state=tk.NORMAL)
        self.btn_atualizar.config(state=tk.DISABLED)
        self.btn_apagar.config(state=tk.DISABLED)
        self.btn_calcular.config(state=tk.NORMAL)            
        self.selected_item = None
        self.entry_search.delete(0, tk.END)

    def clear_fields_2(self): # outra limpeza, porem essa limpa todos os campos , inclusive o treeview - botao Limpar/Esc
        self.entry_data.delete(0, tk.END)
        self.entry_tipo_impressora.config(state='normal')        
        self.entry_tipo_impressora.delete(0, tk.END)
        self.entry_tipo_impressora.config(state='readonly')        
        self.entry_funcionario.config(state='normal')
        self.entry_funcionario.delete(0, tk.END)
        self.entry_funcionario.config(state='readonly')        
        self.entry_cliente.delete(0, tk.END)
        self.entry_n_of.delete(0, tk.END)
        self.entry_quant_cores.delete(0, tk.END)
        self.entry_medida.delete(0, tk.END)
        self.entry_corte.delete(0, tk.END)
        self.entry_quant_papel.delete(0, tk.END)
        self.entry_metro_quadrado.config(state='normal')
        self.entry_metro_quadrado.delete(0, tk.END)
        self.entry_metro_quadrado.config(state='disable') 
        self.entry_tipo_papel.config(state='normal')                  
        self.entry_tipo_papel.delete(0, tk.END)
        self.entry_tipo_papel.config(state='readonly')
        self.entry_horas_rum.delete(0, tk.END)
        self.entry_ta_h.delete(0, tk.END)        
        self.entry_ta_i.delete(0, tk.END) 
        self.entry_hi_a.delete(0, tk.END)         
        self.entry_hi_b.delete(0, tk.END)
        self.entry_hi_c.delete(0, tk.END)
        self.entry_hi_d.delete(0, tk.END) 
        self.entry_hi_e.delete(0, tk.END)         
        self.entry_hi_f.delete(0, tk.END)         
        self.entry_hi_g.delete(0, tk.END)
        self.entry_total_hp.config(state='normal')        
        self.entry_total_hp.delete(0, tk.END)
        self.entry_total_hp.config(state='disable')        
        self.entry_data_inicial.delete(0, tk.END)
        self.entry_data_final.delete(0, tk.END)

        self.btn_cadastrar.config(state=tk.NORMAL)
        self.btn_atualizar.config(state=tk.DISABLED)
        self.btn_apagar.config(state=tk.DISABLED)
        self.btn_calcular.config(state=tk.NORMAL)            
        self.selected_item = None
        self.entry_search.delete(0, tk.END)

        self.tree.delete(*self.tree.get_children())        
        # Limpar o somatório
        self.somatorio_horas_rum_label.config(text="H.Rum: ")
        self.somatorio_quant_papel_label.config(text="Q.Papel: ")
        self.somatorio_ta_h_label.config(text="TA (H): ")
        self.somatorio_ta_i_label.config(text="TA (I): ") 
        self.somatorio_hi_a_label.config(text="A: ")
        self.somatorio_hi_b_label.config(text="B: ")
        self.somatorio_hi_c_label.config(text="C: ")
        self.somatorio_hi_d_label.config(text="D: ")
        self.somatorio_hi_e_label.config(text="E: ")
        self.somatorio_hi_f_label.config(text="F: ")
        self.somatorio_hi_g_label.config(text="G: ")
        self.somatorio_total_hp_label.config(text="Total Horas Paga: ")

        
    def search_data(self):
        search_text = self.entry_search.get().strip().upper()
        data_inicial = self.entry_data_inicial.get_date().strftime('%Y-%m-%d') if self.entry_data_inicial.get() else None
        data_final = self.entry_data_final.get_date().strftime('%Y-%m-%d') if self.entry_data_final.get() else None

        # Constrói a query dinamicamente com base nos filtros disponíveis
        query = '''
            SELECT * FROM RELAPROD
            WHERE (impressora LIKE ? OR funcionario LIKE ? OR cliente LIKE ? OR medida LIKE ? OR tipo_papel LIKE ? OR n_of LIKE ?)
        '''
        
        params = ['%' + search_text + '%'] * 6  # Parâmetros para a pesquisa por texto

        if data_inicial and data_final:
            query += ' AND data BETWEEN ? AND ?'
            params.extend([data_inicial, data_final])

        # Executa a query com os parâmetros corretos
        self.cursor.execute(query, tuple(params))
        rows = self.cursor.fetchall()

        # Atualiza o Treeview com os resultados
        self.tree.delete(*self.tree.get_children())
        for row in rows:
            self.tree.insert("", tk.END, values=row)

        # Atualiza o somatório com base nos itens exibidos no Treeview
        self.update_somatorio()

        # Limpa os campos de data após a pesquisa
        #if data_inicial and data_final:
        #    self.entry_data_inicial.delete(0, tk.END)
        #    self.entry_data_final.delete(0, tk.END)


    def on_double_click(self, event):
        if self.tree.selection():   #testa se foi selecionado um item, caso tenha clicado no menu do treeview
            item = self.tree.selection()[0]
            self.selected_item = self.tree.item(item)

            self.entry_data.config(state='normal')
            self.entry_tipo_impressora.config(state='normal')
            self.entry_funcionario.config(state='normal')
            self.entry_tipo_papel.config(state='normal')            
            # Habilita o campo ao duplo clica para deixar atualizar a caixa de entrada, entry
            self.entry_metro_quadrado.config(state='normal')
            self.entry_total_hp.config(state='normal')                        

            # Limpa os campos de entrada
            self.entry_data.delete(0, tk.END)
            self.entry_tipo_impressora.delete(0, tk.END)            
            self.entry_funcionario.delete(0, tk.END)            
            self.entry_cliente.delete(0, tk.END)
            self.entry_n_of.delete(0, tk.END)
            self.entry_quant_cores.delete(0, tk.END)
            self.entry_medida.delete(0, tk.END)
            self.entry_corte.delete(0, tk.END)
            self.entry_quant_papel.delete(0, tk.END)
            self.entry_metro_quadrado.delete(0, tk.END)            
            self.entry_tipo_papel.delete(0, tk.END)
            self.entry_horas_rum.delete(0, tk.END)
            self.entry_ta_h.delete(0, tk.END)
            self.entry_ta_i.delete(0, tk.END)
            self.entry_hi_a.delete(0, tk.END)            
            self.entry_hi_b.delete(0, tk.END)
            self.entry_hi_c.delete(0, tk.END) 
            self.entry_hi_d.delete(0, tk.END)      
            self.entry_hi_e.delete(0, tk.END)
            self.entry_hi_f.delete(0, tk.END)      
            self.entry_hi_g.delete(0, tk.END)
            self.entry_total_hp.delete(0, tk.END)            

            # Insere os valores nos campos de entrada
            self.entry_data.insert(0, self.selected_item['values'][1])
            self.entry_tipo_impressora.insert(0, self.selected_item['values'][2])            
            self.entry_funcionario.insert(0, self.selected_item['values'][3])
            self.entry_cliente.insert(0, self.selected_item['values'][4])
            self.entry_n_of.insert(0, self.selected_item['values'][5])
            self.entry_quant_cores.insert(0, self.selected_item['values'][6])
            self.entry_medida.insert(0, self.selected_item['values'][7])
            self.entry_corte.insert(0, self.selected_item['values'][8])
            self.entry_quant_papel.insert(0, self.selected_item['values'][9])
            self.entry_metro_quadrado.insert(0, self.selected_item['values'][10])
            self.entry_tipo_papel.insert(0, self.selected_item['values'][11])
            self.entry_horas_rum.insert(0, self.selected_item['values'][12])
            self.entry_ta_h.insert(0, self.selected_item['values'][13])            
            self.entry_ta_i.insert(0, self.selected_item['values'][14])
            self.entry_hi_a.insert(0, self.selected_item['values'][15])
            self.entry_hi_b.insert(0, self.selected_item['values'][16])
            self.entry_hi_c.insert(0, self.selected_item['values'][17])
            self.entry_hi_d.insert(0, self.selected_item['values'][18])            
            self.entry_hi_e.insert(0, self.selected_item['values'][19]) 
            self.entry_hi_f.insert(0, self.selected_item['values'][20])             
            self.entry_hi_g.insert(0, self.selected_item['values'][21])
            self.entry_total_hp.insert(0, self.selected_item['values'][22])

            # ao final da regra, bloqueia caixa de entrada novamente por segurança, apos o update do duplo clique
            self.entry_metro_quadrado.config(state='disable')       
            self.entry_total_hp.config(state='disable')               
            self.entry_data.config(state='readonly')
            self.entry_funcionario.config(state='readonly')
            self.entry_tipo_papel.config(state='readonly')
            self.entry_tipo_impressora.config(state='readonly')            

            self.btn_cadastrar.config(state=tk.DISABLED)
            self.btn_atualizar.config(state=tk.NORMAL)
            self.btn_apagar.config(state=tk.NORMAL)
            self.btn_calcular.config(state=tk.DISABLED)            
        else:
            messagebox.showerror("Erro", "Você clicou no menu.\nNenhum item selecionado.\nSelecione um item da lista,\npara editar e atualizar")

    def amostrar_tudo(self): 
        self.fetch_data()
        self.update_somatorio()

if __name__ == "__main__":
    root = tk.Tk()
    app = ProductManagementApp(root)
    root.mainloop()