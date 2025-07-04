import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import openpyxl 
from datetime import timedelta
from PIL import Image, ImageTk
import mysql.connector
from mysql.connector import Error
import os
import sys
import mysql.connector




def clean_value(value):
    if pd.isna(value) or value == '':
        return 0.0
    if isinstance(value, timedelta):
        hours = int(value.total_seconds() // 3600)
        minutes = int((value.total_seconds() % 3600) // 60)
        return f"{hours}.{minutes:02d}"
    if isinstance(value, str):
        value = value.strip()
        if ":" in value:
            try:
                parts = value.split(":")
                hours = int(parts[0])
                minutes = int(parts[1])
                return f"{hours}.{minutes:02d}"
            except:
                return 0.0
        if value.startswith('R$'):
            try:
                value = value.replace('R$', '').replace('.', '').replace(',', '.').strip()
                return round(float(value), 2)
            except:
                return 0.0
        if value.endswith('%'):
            return float(value.replace('%', ''))
        try:
            return round(float(value), 2)
        except:
            return 0.0
    try:
        return round(float(value), 2)
    except:
        return 0.0

def process_payroll_excel(file_path, empresa, mes, ano, column_to_event):
    df = pd.read_excel(file_path, header=None)
    insert_statements = []
    preview_data = []
    for idx in range(1, len(df)):
        row = df.iloc[idx]
        if pd.notna(row[1]) and str(row[1]).strip().isdigit():
            matricula = int(row[1])
            for col, event_code in column_to_event.items():
                value = row[col]
                if pd.isna(value) or value == '' or value == '-':
                    continue
                cleaned_value = clean_value(value)
                if cleaned_value != 0.0:
                    insert_statements.append(
                        f"""INSERT INTO movevento 
                        (cd_empresa, mes, ano, cd_funcionario, cd_evento, referencia, transferido, tipo_processamento, origem_digitacao)
                        VALUES ({empresa}, {mes}, {ano}, {matricula}, {event_code}, {cleaned_value}, '', 2, 'M')
                        ON DUPLICATE KEY UPDATE 
                            referencia = {cleaned_value},
                            transferido = '',
                            tipo_processamento = 2,
                            origem_digitacao = 'M';"""
                    )
                    preview_data.append((matricula, event_code, cleaned_value))
    return insert_statements, preview_data

def criar_conexao_mysql():
    """Cria conexão MySQL com configurações específicas para evitar erros de localização"""
    try:
        # Configurações específicas para evitar problemas de localização
        config = {
            'host': "192.168.0.2",
            'user': "root",
            'password': "root",
            'database': "ebs_cordilheira",
            'port': 5003,
            'connection_timeout': 10,
            'charset': 'utf8mb4',
            'collation': 'utf8mb4_unicode_ci',
            'use_unicode': True,
            'autocommit': False,
            # Estas configurações ajudam a evitar problemas de localização
            'sql_mode': 'TRADITIONAL',
            'init_command': "SET sql_mode='STRICT_TRANS_TABLES'"
        }
        
        # Tentar conexão com configurações específicas
        conexao = mysql.connector.connect(**config)
        
        # Configurar encoding da sessão
        if conexao.is_connected():
            cursor = conexao.cursor()
            cursor.execute("SET NAMES utf8mb4 COLLATE utf8mb4_unicode_ci")
            cursor.execute("SET CHARACTER SET utf8mb4")
            cursor.execute("SET character_set_connection=utf8mb4")
            cursor.close()
            
        return conexao
        
    except Error as e:
        print(f"Erro MySQL específico: {e}")
        # Tentar conexão mais simples como fallback
        try:
            conexao_simples = mysql.connector.connect(
                host="192.168.0.2",
                user="root",
                password="root",
                database="ebs_cordilheira",
                port=5003,
                connection_timeout=10
            )
            return conexao_simples
        except Error as e2:
            print(f"Erro na conexão fallback: {e2}")
            raise e
    except Exception as e:
        print(f"Erro geral: {e}")
        raise e

def abrir_interface():
    janela = tk.Tk()
    janela.title("Gerador de SQL da Folha de Pagamento")
    janela.geometry("900x620")

    frame_esquerdo = tk.Frame(janela, bg="#0C2238", width=300)
    frame_esquerdo.pack(side="left", fill="y")

    canvas_direito = tk.Canvas(janela, bg="white")
    canvas_direito.pack(side="right", fill="both", expand=True)

    scrollbar_vertical = tk.Scrollbar(janela, orient="vertical", command=canvas_direito.yview)
    scrollbar_vertical.pack(side="right", fill="y")
    canvas_direito.configure(yscrollcommand=scrollbar_vertical.set)

    frame_direito = tk.Frame(canvas_direito, bg="white")
    canvas_direito.create_window((0,0), window=frame_direito, anchor="nw")

    def on_frame_configure(event):
        canvas_direito.configure(scrollregion=canvas_direito.bbox("all"))
    frame_direito.bind("<Configure>", on_frame_configure)

    def _on_mousewheel(event):
        canvas_direito.yview_scroll(int(-1*(event.delta/120)), "units")
    canvas_direito.bind_all("<MouseWheel>", _on_mousewheel)

    try:
        # Verificar se está executando como executável
        if getattr(sys, 'frozen', False):
            # Executando como executável
            base_path = sys._MEIPASS
        else:
            # Executando como script
            base_path = os.path.dirname(os.path.abspath(__file__))
            
        logo_path = os.path.join(base_path, "logo.png")
        
        if os.path.exists(logo_path):
            imagem_logo = Image.open(logo_path)
            imagem_logo = imagem_logo.resize((230, 40))
            logo_tk = ImageTk.PhotoImage(imagem_logo)
            tk.Label(frame_esquerdo, bg="#0C2238").pack(pady=70)
            tk.Label(frame_esquerdo, image=logo_tk, bg="#0C2238").pack()
            tk.Label(frame_esquerdo, bg="#0C2238").pack(pady=15)
        else:
            print(f"Logo não encontrada em: {logo_path}")
    except Exception as e:
        print(f"Erro ao carregar logo: {e}")

    tk.Label(frame_esquerdo, text="Bem-vindo\nGerador de DP", bg="#0C2238", fg="white", font=("Segoe UI Semibold", 18)).pack(pady=(0,5))
    tk.Label(frame_esquerdo, text="Carregue uma planilha,\nassocie colunas e eventos\n", bg="#0C2238", fg="white", font=("Segoe UI", 12), justify="center").pack(pady=(0,20))
    tk.Label(frame_esquerdo, text="Desenvolvido pelo Setor de Inteligência de Dados\n e Automação", 
             bg="#0C2238", fg="white", font=("Segoe UI", 9)).pack(side="bottom", pady=10)

    proxima_coluna = [2]
    colunas_livres = []
    campos_evento = []
    comandos_sql = []
    preview_data_completo = []

    def testar_conexao_ui():
        """Testa a conexão com interface gráfica usando a nova função"""
        try:
            print("Tentando conectar ao banco...")
            conexao = criar_conexao_mysql()
            
            if conexao.is_connected():
                cursor = conexao.cursor()
                
                # Obter informações do servidor de forma mais segura
                cursor.execute("SELECT VERSION()")
                versao_resultado = cursor.fetchone()
                db_info = versao_resultado[0] if versao_resultado else "Desconhecida"
                
                print(f"Conectado com sucesso ao MySQL Server versão {db_info}")
                
                cursor.execute("SELECT DATABASE()")
                nome_db = cursor.fetchone()
                print(f"Conectado ao banco: {nome_db[0]}")
                
                # Testa se a tabela existe
                cursor.execute("SHOW TABLES LIKE 'movevento'")
                tabela_existe = cursor.fetchone()
                if tabela_existe:
                    print("✓ Tabela 'movevento' encontrada!")
                    cursor.execute("SELECT COUNT(*) FROM movevento")
                    total_registros = cursor.fetchone()[0]
                    mensagem = f"Conexão bem-sucedida!\n\nBanco: {nome_db[0]}\nServidor: {db_info}\nTabela 'movevento' encontrada com {total_registros} registros."
                else:
                    mensagem = f"Conexão bem-sucedida!\n\nBanco: {nome_db[0]}\nServidor: {db_info}\n⚠ Tabela 'movevento' não encontrada!"
                    
                cursor.close()
                conexao.close()
                messagebox.showinfo("Teste de Conexão", mensagem)
                return True
                
        except Error as e:
            erro_msg = f"Erro ao conectar ao MySQL:\n{str(e)}"
            print(erro_msg)
            messagebox.showerror("Erro de Conexão", erro_msg)
            return False
        except Exception as e:
            erro_msg = f"Erro inesperado:\n{str(e)}"
            print(erro_msg)
            messagebox.showerror("Erro de Conexão", erro_msg)
            return False

    def selecionar_arquivo():
        caminho = filedialog.askopenfilename(filetypes=[("Planilhas Excel", "*.xlsx *.xls *.ods")])
        if caminho:
            entrada_arquivo.delete(0, tk.END)
            entrada_arquivo.insert(0, caminho)

    def adicionar_linha_evento():
        linha_tela = len(campos_evento) + 7
        if colunas_livres:
            coluna_excel = colunas_livres.pop(0)
        else:
            coluna_excel = proxima_coluna[0]
            proxima_coluna[0] += 1

        coluna_entry = tk.Entry(frame_direito, font=("Segoe UI Semibold", 12), width=10)
        coluna_entry.insert(0, str(coluna_excel))
        coluna_entry.configure(state='readonly')
        coluna_entry.grid(row=linha_tela, column=0, padx=5, pady=2)

        evento_entry = tk.Entry(frame_direito, font=("Segoe UI Semibold", 12), width=10)
        evento_entry.grid(row=linha_tela, column=1, padx=5, pady=2)

        def limpar_campo():
            evento_entry.delete(0, tk.END)

        btn_limpar = ttk.Button(frame_direito, text="🪝", style="Azul.TButton", command=limpar_campo)
        btn_limpar.grid(row=linha_tela, column=2, padx=5)

        campos_evento.append((coluna_excel, evento_entry, coluna_entry, evento_entry, btn_limpar))

    def aplicar_filtro_funcionario(selecionados):
        tabela_preview.delete(*tabela_preview.get_children())
        for matricula, evento, valor in preview_data_completo:
            if matricula in selecionados:
                tabela_preview.insert("", "end", values=(matricula, evento, valor))

    def abrir_janela_filtro():
        if not preview_data_completo:
            messagebox.showwarning("Aviso", "Nenhum dado carregado para filtrar.")
            return

        funcionarios_unicos = sorted(set(m for m, _, _ in preview_data_completo))

        popup = tk.Toplevel(janela)
        popup.title("Filtrar por Funcionário")
        popup.geometry("350x450")
        popup.resizable(False, False)
        popup.grab_set()

        # Centralizar a janela popup
        popup.transient(janela)
        popup.update_idletasks()
        x = (popup.winfo_screenwidth() // 2) - (350 // 2)
        y = (popup.winfo_screenheight() // 2) - (450 // 2)
        popup.geometry(f"350x450+{x}+{y}")

        # Cabeçalho
        header_frame = tk.Frame(popup, bg="#0C2238", height=60)
        header_frame.pack(fill="x", pady=(0, 10))
        header_frame.pack_propagate(False)
        
        tk.Label(header_frame, text="🔍 Filtrar por Funcionário", 
                font=("Segoe UI Semibold", 14), bg="#0C2238", fg="white").pack(expand=True)

        # Frame para seleção de todos/nenhum
        selection_frame = tk.Frame(popup)
        selection_frame.pack(fill="x", padx=10, pady=(0, 10))

        def selecionar_todos():
            for var in check_vars.values():
                var.set(True)

        def desselecionar_todos():
            for var in check_vars.values():
                var.set(False)

        ttk.Button(selection_frame, text="✓ Selecionar Todos", 
                  command=selecionar_todos).pack(side="left", padx=(0, 5))
        ttk.Button(selection_frame, text="✗ Desselecionar Todos", 
                  command=desselecionar_todos).pack(side="left")

        # Frame principal com scroll
        main_frame = tk.Frame(popup)
        main_frame.pack(fill="both", expand=True, padx=10)

        # Canvas e scrollbar para a lista de funcionários
        canvas_scroll = tk.Canvas(main_frame, bg="white", highlightthickness=0)
        scrollbar_popup = tk.Scrollbar(main_frame, orient="vertical", command=canvas_scroll.yview)
        scrollable_frame = tk.Frame(canvas_scroll, bg="white")

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas_scroll.configure(scrollregion=canvas_scroll.bbox("all"))
        )

        canvas_scroll.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas_scroll.configure(yscrollcommand=scrollbar_popup.set)

        # Função para scroll com mouse wheel
        def on_mousewheel_popup(event):
            canvas_scroll.yview_scroll(int(-1*(event.delta/120)), "units")

        # Bind do mouse wheel para todos os widgets relevantes
        def bind_mousewheel(widget):
            widget.bind("<MouseWheel>", on_mousewheel_popup)

        bind_mousewheel(canvas_scroll)
        bind_mousewheel(scrollable_frame)

        # Criar checkboxes para cada funcionário
        check_vars = {}
        for i, funcionario in enumerate(funcionarios_unicos):
            var = tk.BooleanVar(value=True)
            
            # Frame para cada checkbox com background alternado
            checkbox_frame = tk.Frame(scrollable_frame, bg="white" if i % 2 == 0 else "#f0f0f0")
            checkbox_frame.pack(fill="x", pady=1)
            
            chk = tk.Checkbutton(
                checkbox_frame, 
                text=f"Funcionário {funcionario}", 
                variable=var, 
                font=("Segoe UI", 11),
                bg="white" if i % 2 == 0 else "#f0f0f0",
                anchor="w"
            )
            chk.pack(fill="x", padx=10, pady=5)
            
            # Bind mousewheel para cada checkbox também
            bind_mousewheel(chk)
            bind_mousewheel(checkbox_frame)
            
            check_vars[funcionario] = var

        # Posicionar canvas e scrollbar
        canvas_scroll.pack(side="left", fill="both", expand=True)
        scrollbar_popup.pack(side="right", fill="y")

        # Frame para os botões
        button_frame = tk.Frame(popup)
        button_frame.pack(fill="x", padx=10, pady=10)

        def aplicar():
            selecionados = [f for f, v in check_vars.items() if v.get()]
            if not selecionados:
                messagebox.showwarning("Aviso", "Selecione pelo menos um funcionário para aplicar o filtro.")
                return
            aplicar_filtro_funcionario(selecionados)
            popup.destroy()

        def mostrar_todos():
            aplicar_filtro_funcionario(funcionarios_unicos)
            popup.destroy()

        # Estilo dos botões
        style_popup = ttk.Style()
        style_popup.configure("Filtro.TButton", 
                            font=("Segoe UI Semibold", 10), 
                            padding=6)

        ttk.Button(button_frame, text="✓ Aplicar Filtro", 
                  style="Filtro.TButton", command=aplicar).pack(side="left", padx=(0, 5))
        ttk.Button(button_frame, text="👁 Mostrar Todos", 
                  style="Filtro.TButton", command=mostrar_todos).pack(side="left", padx=(0, 5))
        ttk.Button(button_frame, text="✗ Cancelar", 
                  style="Filtro.TButton", command=popup.destroy).pack(side="right")

        # Informações na parte inferior
        info_frame = tk.Frame(popup, bg="#f8f9fa")
        info_frame.pack(fill="x", pady=(5, 0))
        tk.Label(info_frame, text=f"Total de funcionários: {len(funcionarios_unicos)}", 
                font=("Segoe UI", 9), bg="#f8f9fa", fg="#666").pack(pady=5)

        # Foco inicial no canvas para permitir scroll imediato
        canvas_scroll.focus_set()

    def gerar_sql():
        nonlocal comandos_sql, preview_data_completo
        tabela_preview.delete(*tabela_preview.get_children())

        file_path = entrada_arquivo.get()
        empresa = entrada_empresa.get()
        mes = entrada_mes.get()
        ano = entrada_ano.get()

        if not (file_path and empresa and mes and ano):
            messagebox.showerror("Erro", "Preencha todos os campos.")
            return

        confirmacao = messagebox.askyesno(
            "Confirmação de Dados",
            f"Os dados de Empresa: {empresa}, Mês: {mes} e Ano: {ano} estão corretos?\n\nClique em 'Sim' para continuar ou 'Não' para cancelar."
        )
        if not confirmacao:
            return

        eventos = {}
        for col, entrada_evt, *_ in campos_evento:
            evt_text = entrada_evt.get().strip()
            if evt_text:
                try:
                    evt = int(evt_text)
                    eventos[col] = evt
                except:
                    messagebox.showwarning("Aviso", f"Evento inválido na coluna {col}: '{evt_text}' — será ignorado.")

        if not eventos:
            messagebox.showerror("Erro", "Nenhum par coluna:evento válido foi informado.")
            return

        try:
            comandos_sql, preview = process_payroll_excel(file_path, int(empresa), int(mes), int(ano), eventos)
            preview_data_completo = preview.copy()

            aplicar_filtro_funcionario([m for m, _, _ in preview_data_completo])  # Mostra tudo inicialmente

            if comandos_sql:
                btn_aplicar_banco.config(state='normal')
                messagebox.showinfo("Pré-visualização completa", f"{len(comandos_sql)} comandos foram gerados e estão prontos para aplicar ao banco.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar: {e}")

    def aplicar_sql_no_banco(sql_commands):
        """Aplicar comandos SQL no banco com melhor tratamento de erros"""
        conexao = None
        try:
            print(f"Conectando ao banco para executar {len(sql_commands)} comandos...")
            
            conexao = criar_conexao_mysql()
            
            cursor = conexao.cursor()
            
            # Executar todos os comandos em uma transação
            comandos_executados = 0
            for i, comando in enumerate(sql_commands):
                try:
                    cursor.execute(comando)
                    comandos_executados += 1
                    
                    # Progresso a cada 100 comandos
                    if (i + 1) % 100 == 0:
                        print(f"Executados {i + 1}/{len(sql_commands)} comandos...")
                        
                except Error as e:
                    print(f"Erro no comando {i + 1}: {e}")
                    print(f"Comando que falhou: {comando[:100]}...")
                    raise e
            
            # Commit da transação
            conexao.commit()
            print(f"✓ Todos os {comandos_executados} comandos executados com sucesso!")
            
            cursor.close()
            return True, f"{comandos_executados} comandos aplicados com sucesso ao banco."
            
        except Error as e:
            if conexao:
                conexao.rollback()
                print("Transação revertida devido ao erro.")
            return False, f"Erro MySQL: {e}"
        except Exception as e:
            if conexao:
                conexao.rollback()
                print("Transação revertida devido ao erro.")
            return False, f"Erro inesperado: {e}"
        finally:
            if conexao and conexao.is_connected():
                conexao.close()
                print("Conexão fechada.")

    def aplicar_ao_banco():
        if not comandos_sql:
            messagebox.showwarning("Aviso", "Nenhum comando foi gerado ainda.")
            return

        # Confirma antes de aplicar
        confirmacao = messagebox.askyesno(
            "Confirmação de Aplicação",
            f"Você tem certeza que deseja aplicar {len(comandos_sql)} comandos ao banco de dados?\n\nEsta operação não pode ser desfeita."
        )
        if not confirmacao:
            return

        sucesso, msg = aplicar_sql_no_banco(comandos_sql)
        if sucesso:
            messagebox.showinfo("Sucesso", msg)
        else:
            messagebox.showerror("Erro", f"Erro ao aplicar comandos: {msg}")

    def redefinir_tudo():
        for _, _, coluna_entry, evento_entry, btn in campos_evento:
            coluna_entry.grid_remove()
            evento_entry.grid_remove()
            btn.grid_remove()
        campos_evento.clear()
        proxima_coluna[0] = 2
        colunas_livres.clear()
        tabela_preview.delete(*tabela_preview.get_children())
        comandos_sql.clear()
        preview_data_completo.clear()
        btn_aplicar_banco.config(state='disabled')

    def limpar_campos_evento():
        for _, evento_entry, *_ in campos_evento:
            evento_entry.delete(0, tk.END)

    def remover_ultima_linha():
        if campos_evento:
            col, evento_entry, coluna_entry, evento_entry, btn = campos_evento.pop()
            coluna_entry.grid_remove()
            evento_entry.grid_remove()
            btn.grid_remove()
            colunas_livres.append(col)
            colunas_livres.sort()

    estilo = ttk.Style()
    estilo.theme_use("default")
    estilo.configure("Azul.TButton", font=("Segoe UI Semibold", 11), padding=8, background="#0C2238", foreground="white", borderwidth=0)
    estilo.map("Azul.TButton", background=[("active", "#145DA0")], foreground=[("active", "white")])
    
    # Estilo para botão discreto da engrenagem
    estilo.configure("Discreto.TButton", font=("Segoe UI", 10), padding=4, background="#f0f0f0", foreground="#666", borderwidth=1)
    estilo.map("Discreto.TButton", background=[("active", "#e0e0e0")], foreground=[("active", "#333")])

    fonte_label = ("Segoe UI Semibold", 12)
    fonte_entry = ("Segoe UI", 12)

    # Botão de teste de conexão discreto no canto superior direito
    btn_teste_conexao = ttk.Button(frame_direito, text="⚙", style="Discreto.TButton", command=testar_conexao_ui)
    btn_teste_conexao.grid(row=0, column=3, padx=5, pady=4, sticky="e")

    tk.Label(frame_direito, text="Arquivo Excel:", font=fonte_label, bg="white").grid(row=0, column=0, sticky='e', padx=5, pady=4)
    entrada_arquivo = tk.Entry(frame_direito, font=fonte_entry, width=40)
    entrada_arquivo.grid(row=0, column=1, padx=5, pady=4)
    ttk.Button(frame_direito, text="📁 Selecionar", style="Azul.TButton", command=selecionar_arquivo).grid(row=0, column=2, padx=5, pady=4)

    tk.Label(frame_direito, text="Empresa:", font=fonte_label, bg="white").grid(row=1, column=0, sticky='e', padx=5, pady=4)
    entrada_empresa = tk.Entry(frame_direito, font=fonte_entry)
    entrada_empresa.insert(0, "2")
    entrada_empresa.grid(row=1, column=1, padx=5, pady=4)

    tk.Label(frame_direito, text="Mês:", font=fonte_label, bg="white").grid(row=2, column=0, sticky='e', padx=5, pady=4)
    entrada_mes = tk.Entry(frame_direito, font=fonte_entry)
    entrada_mes.insert(0, "6")
    entrada_mes.grid(row=2, column=1, padx=5, pady=4)

    tk.Label(frame_direito, text="Ano:", font=fonte_label, bg="white").grid(row=3, column=0, sticky='e', padx=5, pady=4)
    entrada_ano = tk.Entry(frame_direito, font=fonte_entry)
    entrada_ano.insert(0, "2025")
    entrada_ano.grid(row=3, column=1, padx=5, pady=4)

    tk.Label(frame_direito, text="Coluna Excel", font=fonte_label, bg="#FCD12A").grid(row=5, column=0, pady=(10,0), padx=5)
    tk.Label(frame_direito, text="Código Evento", font=fonte_label, bg="#FCD12A").grid(row=5, column=1, pady=(10,0), padx=5)

    btn_adicionar = ttk.Button(frame_direito, text="➕ Adicionar Linha", style="Azul.TButton", command=adicionar_linha_evento)
    btn_adicionar.grid(row=6, column=2, padx=10, pady=10, sticky="w")

    frame_tabela = tk.Frame(frame_direito, bg="white")
    frame_tabela.grid(row=102, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")

    scrollbar = tk.Scrollbar(frame_tabela)
    scrollbar.pack(side="right", fill="y")

    tabela_preview = ttk.Treeview(frame_tabela, columns=("Funcionario", "Evento", "Valor"), show="headings", height=8, yscrollcommand=scrollbar.set)
    tabela_preview.heading("Funcionario", text="Funcionário")
    tabela_preview.heading("Evento", text="Evento")
    tabela_preview.heading("Valor", text="Valor")
    tabela_preview.pack(side="left", fill="both", expand=True)
    scrollbar.config(command=tabela_preview.yview)

    style_tree = ttk.Style()
    style_tree.configure("Treeview.Heading", background="#FCD12A", font=("Segoe UI Semibold", 12))

    frame_botoes = tk.Frame(frame_direito, bg="white")
    frame_botoes.grid(row=200, column=0, columnspan=3, pady=20)


    # Primeira linha de botões
    ttk.Button(frame_botoes, text="📂 Gerar", style="Azul.TButton", command=gerar_sql).grid(row=0, column=0, padx=8)
    ttk.Button(frame_botoes, text="🔄 Redefinir", style="Azul.TButton", command=redefinir_tudo).grid(row=0, column=1, padx=8)
    ttk.Button(frame_botoes, text="🪝 Limpar", style="Azul.TButton", command=limpar_campos_evento).grid(row=0, column=2, padx=8)
    ttk.Button(frame_botoes, text="➖ Remover Última", style="Azul.TButton", command=remover_ultima_linha).grid(row=0, column=3, padx=8)

    btn_filtro = ttk.Button(frame_botoes, text="🔍 Filtro", style="Azul.TButton", command=abrir_janela_filtro)
    btn_filtro.grid(row=0, column=4, padx=8)

    btn_aplicar_banco = ttk.Button(frame_botoes, text="📤 Aplicar ao Banco", style="Azul.TButton", command=aplicar_ao_banco)
    btn_aplicar_banco.grid(row=0, column=5, padx=8)
    btn_aplicar_banco.config(state='disabled')

    

    frame_direito.grid_rowconfigure(102, weight=1)
    frame_direito.grid_columnconfigure(1, weight=1)

    janela.mainloop()

# Executa a interface
if __name__ == "__main__":
    abrir_interface()