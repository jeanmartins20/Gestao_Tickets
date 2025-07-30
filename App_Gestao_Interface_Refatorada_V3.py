import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
from datetime import datetime, timedelta
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from contextlib import closing
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class DatabaseManager:
    def __init__(self, db_name='tickets.db'):
        self.db_name = db_name
        self._create_table()

    def _create_table(self):
        query = (
            "CREATE TABLE IF NOT EXISTS registros ("
            "id INTEGER PRIMARY KEY AUTOINCREMENT, "
            "data TEXT NOT NULL, "
            "numero_ticket TEXT NOT NULL, "
            "descricao TEXT NOT NULL, "
            "acao_realizada TEXT, "
            "status TEXT)"
        )
        self._execute_query(query)

    def conectar(self):
        return sqlite3.connect(self.db_name)

    def _execute_query(self, query, params=(), fetch=None):
        try:
            with self.conectar() as conn:
                with closing(conn.cursor()) as cursor:
                    cursor.execute(query, params)
                    conn.commit()
                    if fetch == 'one':
                        return cursor.fetchone()
                    if fetch == 'all':
                        return cursor.fetchall()
        except sqlite3.Error as e:
            messagebox.showerror("Erro de Banco de Dados", f"Ocorreu um erro: {e}")
            return None

    def add_record(self, data, numero, descricao, acao, status):
        query = "INSERT INTO registros (data, numero_ticket, descricao, acao_realizada, status) VALUES (?, ?, ?, ?, ?)"
        self._execute_query(query, (data, numero, descricao, acao, status))

    def update_record(self, record_id, data, numero, descricao, acao, status):
        query = "UPDATE registros SET data=?, numero_ticket=?, descricao=?, acao_realizada=?, status=? WHERE id=?"
        self._execute_query(query, (data, numero, descricao, acao, status, record_id))

    def delete_record(self, record_id):
        query = "DELETE FROM registros WHERE id=?"
        self._execute_query(query, (record_id,))

    def fetch_all_records(self):
        return self._execute_query("SELECT * FROM registros ORDER BY id DESC", fetch='all')

    def search_by_number(self, numero):
        query = "SELECT * FROM registros WHERE numero_ticket LIKE ? ORDER BY id DESC"
        return self._execute_query(query, (f'%{numero}%',), fetch='all')

    def fetch_record_by_ticket_number(self, numero_ticket):
        """Busca um registro pelo número do ticket exato."""
        query = "SELECT * FROM registros WHERE numero_ticket = ?"
        return self._execute_query(query, (numero_ticket,), fetch='one')


class TicketApp:
    def __init__(self, root_window):
        self.root = root_window
        self.root.title("Gestão de Tickets de Suporte")
        self.root.state("zoomed")  # Tela cheia
        self.db = DatabaseManager('tickets.db')

        # Dicionário para armazenar as referências dos labels dos balões de estatísticas
        self.stats_labels = {}

        # Configurações de estilo para os balões de estatísticas
        self._configure_styles()

        self._create_widgets()
        self._load_table() # Carrega a tabela e as estatísticas iniciais

    def _configure_styles(self):
        """Configura os estilos para os balões de estatísticas."""
        self.root.style = ttk.Style()
        self.root.style.theme_use('clam') # Um tema que permite mais customização
        self.root.style.configure('blue.TFrame', background='#3498db', relief="solid", borderwidth=1) # Azul
        self.root.style.configure('green.TFrame', background='#2ecc71', relief="solid", borderwidth=1) # Verde
        self.root.style.configure('orange.TFrame', background='#e67e22', relief="solid", borderwidth=1) # Laranja
        self.root.style.configure('purple.TFrame', background='#9b59b6', relief="solid", borderwidth=1) # Roxo

        # Estilos para os labels dentro dos balões
        self.root.style.configure('Stats.TLabel', font=('Segoe UI', 10, 'bold'), foreground="white")
        self.root.style.configure('StatsValue.TLabel', font=('Segoe UI', 18, 'bold'), foreground="white")


    def _create_widgets(self):
        """Cria e organiza todos os widgets da interface principal."""

        # --- Seção Superior: Detalhes do Ticket ---
        input_frame = ttk.LabelFrame(self.root, text="Detalhes do Ticket", padding=15)
        input_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        self.root.columnconfigure(0, weight=1) # Permite que o input_frame se expanda horizontalmente

        # Configuração de colunas para o input_frame
        input_frame.columnconfigure(1, weight=1)
        input_frame.columnconfigure(3, weight=1)
        input_frame.columnconfigure(5, weight=1)

        ttk.Label(input_frame, text="ID:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.id_entry = ttk.Entry(input_frame, width=10, state="readonly")
        self.id_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(input_frame, text="Data (DD/MM/AAAA):").grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.data_entry = ttk.Entry(input_frame)
        self.data_entry.grid(row=0, column=3, padx=5, pady=5, sticky="ew")
        self.data_entry.insert(0, datetime.now().strftime("%d/%m/%Y"))

        ttk.Label(input_frame, text="Nº Ticket/Chamado:").grid(row=0, column=4, padx=5, pady=5, sticky="w")
        self.numero_entry = ttk.Entry(input_frame)
        self.numero_entry.grid(row=0, column=5, padx=5, pady=5, sticky="ew")

        ttk.Label(input_frame, text="Descrição:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.descricao_entry = ttk.Entry(input_frame)
        self.descricao_entry.grid(row=1, column=1, columnspan=5, padx=5, pady=5, sticky="ew")

        ttk.Label(input_frame, text="Ação Realizada:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.acao_entry = ttk.Entry(input_frame)
        self.acao_entry.grid(row=2, column=1, columnspan=3, padx=5, pady=5, sticky="ew")

        ttk.Label(input_frame, text="Status:").grid(row=2, column=4, padx=5, pady=5, sticky="w")
        # Status atualizados em ordem alfabética
        status_options = ["Aguardando Parceiro", "Cancelado", "Em Andamento", "Fechado", "Pendente de Resposta", "Resolvido"]
        self.status_combobox = ttk.Combobox(input_frame, values=status_options, state="readonly")
        self.status_combobox.grid(row=2, column=5, padx=5, pady=5, sticky="ew")
        self.status_combobox.set("Em Andamento") # Define um status inicial

        # --- Seção do Meio: Botões de Ação e Tabela de Tickets ---
        # Frame para os botões de ação
        button_frame = ttk.Frame(self.root, padding=10)
        button_frame.grid(row=1, column=0, padx=10, sticky="ew") # Posicionado abaixo do input_frame

        ttk.Button(button_frame, text="Adicionar", command=self._add_record).pack(side="left", padx=5)
        # Botão Editar agora chama a nova janela de edição
        ttk.Button(button_frame, text="Editar", command=self._open_edit_window).pack(side="left", padx=5)
        # Botão Deletar agora chama a nova janela de deleção
        ttk.Button(button_frame, text="Deletar", command=self._open_delete_window).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Buscar por Nº", command=self._search_record).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Limpar Campos", command=self._clear_fields).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Resumo e Gráfico", command=self._show_chart_popup).pack(side="left", padx=5) # Renomeado para clareza
        ttk.Button(button_frame, text="Importar Dados", command=self._import_data).pack(side="left", padx=5)


        # Frame para a Treeview (tabela de tickets)
        tree_frame = ttk.Frame(self.root)
        tree_frame.grid(row=2, column=0, padx=10, pady=10, sticky="nsew") # Posicionado abaixo dos botões
        self.root.rowconfigure(2, weight=1) # Permite que a tree_frame se expanda verticalmente

        cols = ("ID", "Data", "Nº Ticket", "Descrição", "Ação Realizada", "Status")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show='headings', height=20)
        for col in cols:
            self.tree.heading(col, text=col)
        self.tree.column("ID", width=50, anchor="center")
        self.tree.column("Data", width=100, anchor="center")
        self.tree.column("Nº Ticket", width=120, anchor="center")
        self.tree.column("Descrição", width=300)
        self.tree.column("Ação Realizada", width=300)
        self.tree.column("Status", width=120, anchor="center")
        self.tree.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.grid(row=0, column=1, sticky="ns")

        self.tree.bind("<<TreeviewSelect>>", self._fill_fields_on_select)

        # --- Seção Inferior: Cards de Estatísticas ---
        self.statistics_frame = ttk.Frame(self.root, padding=10)
        self.statistics_frame.grid(row=3, column=0, padx=10, pady=10, sticky="ew") # Posicionado abaixo da tree_frame
        self.statistics_frame.columnconfigure(0, weight=1)
        self.statistics_frame.columnconfigure(1, weight=1)
        self.statistics_frame.columnconfigure(2, weight=1)
        self.statistics_frame.columnconfigure(3, weight=1)

        self._create_statistics_cards(self.statistics_frame)


    def _create_statistics_cards(self, parent_frame):
        """Cria os widgets dos balões de estatísticas na parte inferior da tela."""
        # Função auxiliar para criar um balão de resumo
        def _create_balloon(parent, text_label, value_key, color):
            frame = ttk.Frame(parent, style=f"{color}.TFrame")
            frame.pack(side="left", padx=5, pady=5, fill="both", expand=True) # expand=True para distribuir igualmente

            # Usamos uma label para o texto fixo do balão
            ttk.Label(frame, text=text_label, style='Stats.TLabel', background=color).pack(pady=2)
            # Usamos outra label para o valor, que será atualizado dinamicamente
            value_label = ttk.Label(frame, text="0", style='StatsValue.TLabel', background=color)
            value_label.pack(pady=2)
            self.stats_labels[value_key] = value_label # Armazena a referência para atualização

        _create_balloon(parent_frame, "Total de Tickets", "total_tickets", "blue")
        _create_balloon(parent_frame, "Tickets Tratados Hoje", "treated_today", "green")
        _create_balloon(parent_frame, "Tickets Tratados Semana", "treated_week", "orange")
        _create_balloon(parent_frame, "Tickets Tratados Mês", "treated_month", "purple")

    def _update_statistics_cards(self):
        """Atualiza os valores exibidos nos balões de estatísticas."""
        records = self.db.fetch_all_records()
        if not records:
            # Se não houver registros, reseta todos os contadores para zero
            for key in self.stats_labels:
                self.stats_labels[key].config(text="0")
            return

        df = pd.DataFrame(records, columns=["id", "data", "numero", "descricao", "acao", "status"])

        # Garantir que a coluna 'data' seja do tipo datetime, tratando erros
        df['data'] = pd.to_datetime(df['data'], format='%d/%m/%Y', errors='coerce')
        df.dropna(subset=['data'], inplace=True) # Remove linhas onde a data não pôde ser convertida

        if df.empty:
            # Se não houver dados válidos após a limpeza, reseta os contadores
            for key in self.stats_labels:
                self.stats_labels[key].config(text="0")
            return

        # Calcular estatísticas
        hoje = datetime.now().date()
        inicio_semana = hoje - timedelta(days=hoje.weekday())
        inicio_mes = hoje.replace(day=1)

        # print(f"Hoje: {hoje}")
        # print(f"Início da Semana: {inicio_semana}")
        # print(f"Início do Mês: {inicio_mes}")
        # print(f"DataFrame antes do filtro de status:\n{df[['data', 'status']].head()}")

        total_tickets = df.shape[0]
        # Filtra tickets tratados (status 'Resolvido' ou 'Fechado')
        df_treated = df[df['status'].isin(['Resolvido', 'Fechado'])]

        # print(f"DataFrame de tickets tratados:\n{df_treated[['data', 'status']].head()}")

        # Certifique-se de comparar apenas a parte da data
        treated_today = df_treated[df_treated['data'].dt.date == hoje].shape[0]
        treated_week = df_treated[df_treated['data'].dt.date >= inicio_semana].shape[0]
        treated_month = df_treated[df_treated['data'].dt.date >= inicio_mes].shape[0]

        # print(f"Tickets Tratados Hoje: {treated_today}")
        # print(f"Tickets Tratados Semana: {treated_week}")
        # print(f"Tickets Tratados Mês: {treated_month}")


        # Atualiza os labels dos balões
        self.stats_labels["total_tickets"].config(text=str(total_tickets))
        self.stats_labels["treated_today"].config(text=str(treated_today))
        self.stats_labels["treated_week"].config(text=str(treated_week))
        self.stats_labels["treated_month"].config(text=str(treated_month))

    def _show_chart_popup(self):
        """Exibe o gráfico de colunas em uma nova janela pop-up."""
        records = self.db.fetch_all_records()
        if not records:
            messagebox.showinfo("Gráfico", "Não há dados para gerar o gráfico.")
            return

        df = pd.DataFrame(records, columns=["id", "data", "numero", "descricao", "acao", "status"])
        df['data'] = pd.to_datetime(df['data'], format='%d/%m/%Y', errors='coerce')
        df.dropna(subset=['data'], inplace=True)

        if df.empty:
            messagebox.showinfo("Gráfico", "Não há dados válidos para gerar o gráfico após a limpeza.")
            return

        # Gerar dados para o gráfico de colunas: Tickets por Status
        status_counts = df['status'].value_counts()
        labels = status_counts.index.tolist()
        values = status_counts.values.tolist()

        # Criar a janela pop-up para o gráfico
        chart_window = tk.Toplevel(self.root)
        chart_window.title("Gráfico de Tickets por Status")
        chart_window.state("zoomed") # Abrir em tela cheia

        # Configurar o gráfico de colunas
        fig_bar, ax_bar = plt.subplots(figsize=(10, 6))
        ax_bar.bar(labels, values, color=['steelblue', 'olivedrab', 'indianred', 'mediumorchid', 'darkorange', 'gray']) # Adicionado mais cores
        ax_bar.set_title("Tickets por Status")
        ax_bar.set_xlabel("Status")
        ax_bar.set_ylabel("Número de Tickets")
        ax_bar.grid(axis='y', linestyle='--', alpha=0.7)
        plt.xticks(rotation=45, ha='right') # Rotacionar rótulos para melhor visualização
        plt.tight_layout()

        # Integrar o gráfico no Tkinter
        canvas_bar = FigureCanvasTkAgg(fig_bar, master=chart_window)
        canvas_widget_bar = canvas_bar.get_tk_widget()
        canvas_widget_bar.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=20, pady=10)
        canvas_bar.draw()

    def _get_valid_date(self):
        date_str = self.data_entry.get().strip()
        try:
            return datetime.strptime(date_str, "%d/%m/%Y").strftime("%d/%m/%Y")
        except ValueError:
            messagebox.showwarning("Data inválida", "A data deve estar no formato DD/MM/AAAA.")
            return None

    def _load_table(self, records=None):
        """Carrega os dados na Treeview e atualiza os balões de estatísticas."""
        for row in self.tree.get_children():
            self.tree.delete(row)
        if records is None:
            records = self.db.fetch_all_records()
        if records:
            for row in records:
                self.tree.insert("", tk.END, values=row)
        self._update_statistics_cards() # Chama a atualização das estatísticas após carregar a tabela

    def _validate_inputs(self):
        # Esta validação é para os campos da tela principal (Adicionar)
        if not self.numero_entry.get() or not self.descricao_entry.get():
            messagebox.showwarning("Campos Vazios", "Os campos 'Nº Ticket/Chamado' e 'Descrição' são obrigatórios.")
            return False
        return True

    def _add_record(self):
        if not self._validate_inputs():
            return
        date = self._get_valid_date()
        if not date:
            return
        self.db.add_record(date, self.numero_entry.get(), self.descricao_entry.get(),
                           self.acao_entry.get(), self.status_combobox.get())
        messagebox.showinfo("Sucesso", "Registro adicionado com sucesso!")
        self._load_table()
        self._clear_fields()

    # --- NOVO: Método para abrir a janela de deleção de múltiplos registros ---
    def _open_delete_window(self):
        delete_window = tk.Toplevel(self.root)
        delete_window.title("Deletar Registros")
        delete_window.transient(self.root) # Define a janela principal como pai
        delete_window.grab_set() # Bloqueia interação com a janela principal
        delete_window.focus_set() # Foca na nova janela
        delete_window.state("normal") # Alterado de "zoomed" para "normal"

        # Frame para a Treeview na janela de deleção
        delete_tree_frame = ttk.Frame(delete_window, padding=10)
        delete_tree_frame.pack(fill="both", expand=True)

        cols = ("ID", "Data", "Nº Ticket", "Descrição", "Ação Realizada", "Status")
        # Permite seleção múltipla
        delete_tree = ttk.Treeview(delete_tree_frame, columns=cols, show='headings', selectmode='extended')
        for col in cols:
            delete_tree.heading(col, text=col)
        delete_tree.column("ID", width=50, anchor="center")
        delete_tree.column("Data", width=100, anchor="center")
        delete_tree.column("Nº Ticket", width=120, anchor="center")
        delete_tree.column("Descrição", width=300)
        delete_tree.column("Ação Realizada", width=300)
        delete_tree.column("Status", width=120, anchor="center")
        delete_tree.pack(side="left", fill="both", expand=True)

        delete_scrollbar = ttk.Scrollbar(delete_tree_frame, orient="vertical", command=delete_tree.yview)
        delete_tree.configure(yscrollcommand=delete_scrollbar.set)
        delete_scrollbar.pack(side="right", fill="y")

        # Carrega todos os registros na Treeview da janela de deleção
        records = self.db.fetch_all_records()
        if records:
            for row in records:
                delete_tree.insert("", tk.END, values=row)

        def perform_delete():
            selected_items = delete_tree.selection()
            if not selected_items:
                messagebox.showwarning("Nenhum Selecionado", "Selecione um ou mais registros para deletar.")
                return

            if messagebox.askyesno("Confirmar Deleção", f"Tem certeza que deseja deletar {len(selected_items)} registro(s) selecionado(s)?"):
                deleted_count = 0
                for item_id in selected_items:
                    record_id = delete_tree.item(item_id, 'values')[0] # Pega o ID do registro
                    self.db.delete_record(record_id)
                    deleted_count += 1
                messagebox.showinfo("Sucesso", f"{deleted_count} registro(s) deletado(s) com sucesso!")
                delete_window.destroy() # Fecha a janela de deleção
                self._load_table() # Recarrega a tabela principal e as estatísticas

        delete_button_frame = ttk.Frame(delete_window, padding=10)
        delete_button_frame.pack(fill="x")
        ttk.Button(delete_button_frame, text="Deletar Selecionados", command=perform_delete).pack(side="left", padx=5)
        ttk.Button(delete_button_frame, text="Cancelar", command=delete_window.destroy).pack(side="right", padx=5)

    # --- NOVO: Método para abrir a janela de edição de registro ---
    def _open_edit_window(self):
        edit_window = tk.Toplevel(self.root)
        edit_window.title("Editar Registro")
        edit_window.transient(self.root)
        edit_window.grab_set()
        edit_window.focus_set()
        edit_window.state("normal") # Alterado de "zoomed" para "normal"

        # Frame para busca e seleção de ticket
        search_frame = ttk.LabelFrame(edit_window, text="Buscar Ticket para Editar", padding=10)
        search_frame.pack(pady=10, padx=10, fill="x")
        search_frame.columnconfigure(1, weight=1)

        ttk.Label(search_frame, text="Nº Ticket/ID:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        search_entry = ttk.Entry(search_frame)
        search_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        # Frame para a Treeview de seleção
        edit_tree_frame = ttk.Frame(edit_window, padding=10)
        edit_tree_frame.pack(fill="both", expand=True)

        cols = ("ID", "Data", "Nº Ticket", "Descrição", "Ação Realizada", "Status")
        edit_tree = ttk.Treeview(edit_tree_frame, columns=cols, show='headings', selectmode='browse') # 'browse' para seleção única
        for col in cols:
            edit_tree.heading(col, text=col)
        edit_tree.column("ID", width=50, anchor="center")
        edit_tree.column("Data", width=100, anchor="center")
        edit_tree.column("Nº Ticket", width=120, anchor="center")
        edit_tree.column("Descrição", width=300)
        edit_tree.column("Ação Realizada", width=300)
        edit_tree.column("Status", width=120, anchor="center")
        edit_tree.pack(side="left", fill="both", expand=True)

        edit_scrollbar = ttk.Scrollbar(edit_tree_frame, orient="vertical", command=edit_tree.yview)
        edit_tree.configure(yscrollcommand=edit_scrollbar.set)
        edit_scrollbar.pack(side="right", fill="y")

        def load_edit_tree(records_to_load=None):
            for row in edit_tree.get_children():
                edit_tree.delete(row)
            records = records_to_load if records_to_load is not None else self.db.fetch_all_records()
            if records:
                for row in records:
                    edit_tree.insert("", tk.END, values=row)
        load_edit_tree() # Carrega todos os registros inicialmente

        # Frame para os campos de edição
        edit_fields_frame = ttk.LabelFrame(edit_window, text="Dados do Ticket Selecionado", padding=15)
        edit_fields_frame.pack(pady=10, padx=10, fill="x")
        edit_fields_frame.columnconfigure(1, weight=1)
        edit_fields_frame.columnconfigure(3, weight=1)
        edit_fields_frame.columnconfigure(5, weight=1)

        # Campos de entrada para edição
        edit_id_entry = ttk.Entry(edit_fields_frame, width=10, state="readonly")
        ttk.Label(edit_fields_frame, text="ID:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        edit_id_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        edit_data_entry = ttk.Entry(edit_fields_frame)
        ttk.Label(edit_fields_frame, text="Data (DD/MM/AAAA):").grid(row=0, column=2, padx=5, pady=5, sticky="w")
        edit_data_entry.grid(row=0, column=3, padx=5, pady=5, sticky="ew")

        edit_numero_entry = ttk.Entry(edit_fields_frame)
        ttk.Label(edit_fields_frame, text="Nº Ticket/Chamado:").grid(row=0, column=4, padx=5, pady=5, sticky="w")
        edit_numero_entry.grid(row=0, column=5, padx=5, pady=5, sticky="ew")

        edit_descricao_entry = ttk.Entry(edit_fields_frame)
        ttk.Label(edit_fields_frame, text="Descrição:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        edit_descricao_entry.grid(row=1, column=1, columnspan=5, padx=5, pady=5, sticky="ew")

        edit_acao_entry = ttk.Entry(edit_fields_frame)
        ttk.Label(edit_fields_frame, text="Ação Realizada:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        edit_acao_entry.grid(row=2, column=1, columnspan=3, padx=5, pady=5, sticky="ew")

        edit_status_options = ["Aguardando Parceiro", "Cancelado", "Em Andamento", "Fechado", "Pendente de Resposta", "Resolvido"]
        edit_status_combobox = ttk.Combobox(edit_fields_frame, values=edit_status_options, state="readonly")
        ttk.Label(edit_fields_frame, text="Status:").grid(row=2, column=4, padx=5, pady=5, sticky="w")
        edit_status_combobox.grid(row=2, column=5, padx=5, pady=5, sticky="ew")


        def fill_edit_fields(values):
            """Preenche os campos de edição com os valores do registro selecionado."""
            edit_id_entry.config(state="normal")
            edit_id_entry.delete(0, tk.END)
            edit_id_entry.insert(0, values[0])
            edit_id_entry.config(state="readonly")

            edit_data_entry.delete(0, tk.END)
            edit_data_entry.insert(0, values[1])

            edit_numero_entry.delete(0, tk.END)
            edit_numero_entry.insert(0, values[2])

            edit_descricao_entry.delete(0, tk.END)
            edit_descricao_entry.insert(0, values[3])

            edit_acao_entry.delete(0, tk.END)
            edit_acao_entry.insert(0, values[4])

            edit_status_combobox.set(values[5])

        def on_edit_tree_select(event):
            selected_item = edit_tree.focus()
            if selected_item:
                values = edit_tree.item(selected_item, 'values')
                fill_edit_fields(values)

        edit_tree.bind("<<TreeviewSelect>>", on_edit_tree_select)

        def search_ticket_for_edit():
            search_term = search_entry.get().strip()
            if not search_term:
                load_edit_tree() # Se vazio, recarrega tudo
                messagebox.showwarning("Busca Vazia", "Por favor, digite o número do ticket ou ID para buscar.")
                return

            # Tenta buscar por ID primeiro, se for numérico
            try:
                record_id = int(search_term)
                record = self.db._execute_query("SELECT * FROM registros WHERE id = ?", (record_id,), fetch='one')
            except ValueError:
                record = None # Não é um ID numérico

            # Se não encontrou por ID, tenta buscar por numero_ticket
            if not record:
                record = self.db.fetch_record_by_ticket_number(search_term)

            if record:
                load_edit_tree([record]) # Carrega apenas o registro encontrado na Treeview
                fill_edit_fields(record) # Preenche os campos de edição
            else:
                messagebox.showinfo("Não Encontrado", f"Nenhum ticket encontrado com o número/ID: {search_term}")
                load_edit_tree() # Recarrega a Treeview completa se não encontrou

        ttk.Button(search_frame, text="Buscar Ticket", command=search_ticket_for_edit).grid(row=0, column=2, padx=5, pady=5, sticky="e")


        def save_edited_record():
            record_id = edit_id_entry.get()
            if not record_id:
                messagebox.showwarning("Nenhum Registro", "Selecione um registro para editar antes de salvar.")
                return

            # Validações básicas
            data = edit_data_entry.get().strip()
            numero = edit_numero_entry.get().strip()
            descricao = edit_descricao_entry.get().strip()
            acao = edit_acao_entry.get().strip()
            status = edit_status_combobox.get().strip()

            if not numero or not descricao or not data:
                messagebox.showwarning("Campos Obrigatórios", "Os campos 'Data', 'Nº Ticket/Chamado' e 'Descrição' são obrigatórios.")
                return

            try:
                # Valida o formato da data
                datetime.strptime(data, "%d/%m/%Y")
            except ValueError:
                messagebox.showwarning("Data Inválida", "A data deve estar no formato DD/MM/AAAA.")
                return

            if messagebox.askyesno("Confirmar Edição", "Tem certeza que deseja salvar as alterações?"):
                self.db.update_record(record_id, data, numero, descricao, acao, status)
                messagebox.showinfo("Sucesso", "Registro atualizado com sucesso!")
                edit_window.destroy() # Fecha a janela de edição
                self._load_table() # Recarrega a tabela principal e as estatísticas

        edit_button_frame = ttk.Frame(edit_window, padding=10)
        edit_button_frame.pack(fill="x")
        ttk.Button(edit_button_frame, text="Salvar Edição", command=save_edited_record).pack(side="left", padx=5)
        ttk.Button(edit_button_frame, text="Cancelar", command=edit_window.destroy).pack(side="right", padx=5)


    def _get_valid_date(self):
        date_str = self.data_entry.get().strip()
        try:
            return datetime.strptime(date_str, "%d/%m/%Y").strftime("%d/%m/%Y")
        except ValueError:
            messagebox.showwarning("Data inválida", "A data deve estar no formato DD/MM/AAAA.")
            return None

    def _load_table(self, records=None):
        """Carrega os dados na Treeview e atualiza os balões de estatísticas."""
        for row in self.tree.get_children():
            self.tree.delete(row)
        if records is None:
            records = self.db.fetch_all_records()
        if records:
            for row in records:
                self.tree.insert("", tk.END, values=row)
        self._update_statistics_cards() # Chama a atualização das estatísticas após carregar a tabela

    def _validate_inputs(self):
        # Esta validação é para os campos da tela principal (Adicionar)
        if not self.numero_entry.get() or not self.descricao_entry.get():
            messagebox.showwarning("Campos Vazios", "Os campos 'Nº Ticket/Chamado' e 'Descrição' são obrigatórios.")
            return False
        return True

    def _search_record(self):
        search_term = self.numero_entry.get()
        if not search_term:
            self._load_table()
        else:
            self._load_table(self.db.search_by_number(search_term))

    def _clear_fields(self):
        self.id_entry.config(state="normal")
        self.id_entry.delete(0, tk.END)
        self.id_entry.config(state="readonly")
        self.data_entry.delete(0, tk.END)
        self.data_entry.insert(0, datetime.now().strftime("%d/%m/%Y"))
        self.numero_entry.delete(0, tk.END)
        self.descricao_entry.delete(0, tk.END)
        self.acao_entry.delete(0, tk.END)
        # Garante que a combobox de status seja atualizada com a nova ordem
        self.status_combobox.set('Em Andamento')
        self.numero_entry.focus()

    def _fill_fields_on_select(self, event):
        selected_item = self.tree.focus()
        if not selected_item:
            return
        values = self.tree.item(selected_item, 'values')
        self._clear_fields()
        self.id_entry.config(state="normal")
        self.id_entry.insert(0, values[0])
        self.id_entry.config(state="readonly")
        self.data_entry.insert(0, values[1])
        self.numero_entry.insert(0, values[2])
        self.descricao_entry.insert(0, values[3])
        self.acao_entry.insert(0, values[4])
        # Garante que a combobox de status seja atualizada com a nova ordem
        self.status_combobox.set(values[5])

    def _import_data(self):
        """
        Permite ao usuário selecionar um arquivo CSV ou Excel e importa os dados para o banco de dados.
        Espera que o arquivo tenha as colunas: 'data', 'numero_ticket', 'descricao', 'acao_realizada', 'status'.
        """
        file_path = filedialog.askopenfilename(
            title="Selecionar Arquivo para Importar",
            filetypes=[("Arquivos CSV", "*.csv"), ("Arquivos Excel", "*.xlsx"), ("Todos os Arquivos", "*.*")]
        )

        if not file_path:
            return # Usuário cancelou a seleção

        try:
            # Determina o tipo de arquivo pela extensão
            if file_path.endswith('.csv'):
                df = pd.read_csv(file_path)
            elif file_path.endswith('.xlsx'):
                df = pd.read_excel(file_path)
            else:
                messagebox.showerror("Erro de Importação", "Formato de arquivo não suportado. Por favor, selecione um arquivo CSV ou Excel.")
                return

            # Define as colunas esperadas no arquivo de importação
            expected_columns = ['data', 'numero_ticket', 'descricao', 'acao_realizada', 'status']

            # Verifica se todas as colunas esperadas estão presentes no DataFrame
            if not all(col in df.columns for col in expected_columns):
                missing_cols = [col for col in expected_columns if col not in df.columns]
                messagebox.showerror(
                    "Erro de Colunas",
                    f"O arquivo importado não contém todas as colunas esperadas. "
                    f"Colunas faltando: {', '.join(missing_cols)}\n"
                    f"Certifique-se de que o arquivo possui as colunas: {', '.join(expected_columns)}."
                )
                return

            # Itera sobre as linhas do DataFrame e adiciona ao banco de dados
            imported_count = 0
            for index, row in df.iterrows():
                try:
                    # Converte a data para o formato DD/MM/AAAA se necessário
                    # pd.to_datetime pode ser útil aqui para flexibilidade
                    data = pd.to_datetime(row['data']).strftime("%d/%m/%Y")
                    self.db.add_record(
                        data,
                        str(row['numero_ticket']), # Garante que seja string
                        str(row['descricao']),
                        str(row.get('acao_realizada', '')), # Usa .get para colunas opcionais
                        str(row.get('status', 'Em Andamento')) # Define um padrão se status não existir
                    )
                    imported_count += 1
                except Exception as e:
                    # Loga ou informa o usuário sobre linhas que falharam
                    print(f"Erro ao importar linha {index+2}: {row.to_dict()} - Erro: {e}")

            messagebox.showinfo("Importação Concluída", f"{imported_count} registros importados com sucesso!")
            self._load_table() # Recarrega a tabela para mostrar os novos dados e atualizar estatísticas

        except Exception as e:
            messagebox.showerror("Erro de Importação", f"Ocorreu um erro ao importar o arquivo: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = TicketApp(root)
    root.mainloop()