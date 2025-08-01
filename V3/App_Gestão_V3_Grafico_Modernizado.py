import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
from datetime import datetime, timedelta
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from contextlib import closing
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
# A linha abaixo para Axes3D não é mais estritamente necessária para gráficos 2D,
# mas mantê-la não causa problemas se não for usada.
# from mpl_toolkits.mplot3d import Axes3D

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
        """Configura os estilos para os balões de estatísticas e balões arredondados."""
        self.root.style = ttk.Style()
        self.root.style.theme_use('clam') # Um tema que permite mais customização

        # Estilos para os balões de estatísticas com bordas arredondadas
        self.root.style.configure('Stats.TFrame', background='#3498db', relief="flat", borderwidth=0, border_radius=10)
        self.root.style.map('Stats.TFrame', background=[('active', '#2980b9')]) # Efeito hover

        self.root.style.configure('blue.Stats.TFrame', background='#3498db', relief="flat", borderwidth=0)
        self.root.style.configure('green.Stats.TFrame', background='#2ecc71', relief="flat", borderwidth=0)
        self.root.style.configure('orange.Stats.TFrame', background='#e67e22', relief="flat", borderwidth=0)
        self.root.style.configure('purple.Stats.TFrame', background='#9b59b6', relief="flat", borderwidth=0)

        # Estilos para os labels dentro dos balões
        self.root.style.configure('Stats.TLabel', font=('Segoe UI', 10), foreground="white", background='#3498db')
        self.root.style.configure('StatsValue.TLabel', font=('Segoe UI', 18, 'bold'), foreground="white", background='#3498db')

        # Estilos específicos para cada cor de balão
        self.root.style.configure('blue.Stats.TLabel', background='#3498db')
        self.root.style.configure('blue.StatsValue.TLabel', background='#3498db')
        self.root.style.configure('green.Stats.TLabel', background='#2ecc71')
        self.root.style.configure('green.StatsValue.TLabel', background='#2ecc71')
        self.root.style.configure('orange.Stats.TLabel', background='#e67e22')
        self.root.style.configure('orange.StatsValue.TLabel', background='#e67e22')
        self.root.style.configure('purple.Stats.TLabel', background='#9b59b6')
        self.root.style.configure('purple.StatsValue.TLabel', background='#9b59b6')


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
        self.data_entry.bind("<FocusOut>", self._validate_date_input) # Valida ao sair do campo

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
        status_options = ["Aguardando Parceiro", "Cancelado", "Em Andamento", "Fechado", "Pendente de Resposta", "Resolvido"]
        self.status_combobox = ttk.Combobox(input_frame, values=status_options, state="readonly")
        self.status_combobox.grid(row=2, column=5, padx=5, pady=5, sticky="ew")
        self.status_combobox.set("Em Andamento") # Define um status inicial

        # --- Seção do Meio: Botões de Ação e Tabela de Tickets ---
        # Frame para os botões de ação
        button_frame = ttk.Frame(self.root, padding=10)
        button_frame.grid(row=1, column=0, padx=10, sticky="ew") # Posicionado abaixo do input_frame

        ttk.Button(button_frame, text="Adicionar", command=self._add_record).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Editar", command=self._open_edit_window).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Deletar", command=self._open_delete_window).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Buscar por Nº", command=self._search_record).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Limpar Campos", command=self._clear_fields).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Resumo e Gráfico", command=self._show_chart_popup).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Importar Dados", command=self._import_data).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Exportar Dados", command=self._export_data).pack(side="left", padx=5)


        # Frame para a Treeview (tabela de tickets)
        tree_frame = ttk.Frame(self.root)
        tree_frame.grid(row=2, column=0, padx=10, pady=10, sticky="nsew") # Posicionado abaixo dos botões
        self.root.rowconfigure(2, weight=1) # Permite que a tree_frame se expanda verticalmente

        # Filtro por Status
        filter_frame = ttk.Frame(tree_frame, padding=5)
        filter_frame.pack(fill="x", pady=(0, 5))
        ttk.Label(filter_frame, text="Filtrar por Status:").pack(side="left", padx=5)
        self.filter_status_combobox = ttk.Combobox(filter_frame, values=["Todos"] + status_options, state="readonly")
        self.filter_status_combobox.set("Todos")
        self.filter_status_combobox.pack(side="left", padx=5)
        self.filter_status_combobox.bind("<<ComboboxSelected>>", self._apply_status_filter)


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
        self.tree.pack(side="left", fill="both", expand=True) # Usa pack em vez de grid para melhor expansão

        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y") # Usa pack em vez de grid

        self.tree.bind("<<TreeviewSelect>>", self._fill_fields_on_select)

        # --- Seção Inferior: Balões de Estatísticas ---
        self.statistics_frame = ttk.Frame(self.root, padding=10)
        self.statistics_frame.grid(row=3, column=0, padx=10, pady=10, sticky="ew")
        self.statistics_frame.columnconfigure(0, weight=1)
        self.statistics_frame.columnconfigure(1, weight=1)
        self.statistics_frame.columnconfigure(2, weight=1)
        self.statistics_frame.columnconfigure(3, weight=1)

        self._create_statistics_balloons(self.statistics_frame)


    def _create_statistics_balloons(self, parent_frame):
        """Cria os widgets dos balões de estatísticas na parte inferior da tela com bordas arredondadas."""
        def _create_balloon(parent, text_label, value_key, color):
            # Cria um frame com o estilo de balão arredondado
            frame = ttk.Frame(parent, style=f"{color}.Stats.TFrame", padding=10)
            frame.pack(side="left", padx=10, pady=5, fill="both", expand=True)

            # Labels para o texto e o valor
            ttk.Label(frame, text=text_label, style=f'{color}.Stats.TLabel').pack(pady=2)
            value_label = ttk.Label(frame, text="0", style=f'{color}.StatsValue.TLabel')
            value_label.pack(pady=2)
            self.stats_labels[value_key] = value_label

        _create_balloon(parent_frame, "Total Tickets", "total_tickets", "blue")
        _create_balloon(parent_frame, "Tratados Hoje", "treated_today", "green")
        _create_balloon(parent_frame, "Tratados Semana", "treated_week", "orange")
        _create_balloon(parent_frame, "Tratados Mês", "treated_month", "purple")

    def _update_statistics_cards(self):
        """Atualiza os valores exibidos nos balões de estatísticas."""
        records = self.db.fetch_all_records()
        if not records:
            for key in self.stats_labels:
                self.stats_labels[key].config(text="0")
            return

        df = pd.DataFrame(records, columns=["id", "data", "numero", "descricao", "acao", "status"])
        df['data'] = pd.to_datetime(df['data'], format='%d/%m/%Y', errors='coerce')
        df.dropna(subset=['data'], inplace=True)

        if df.empty:
            for key in self.stats_labels:
                self.stats_labels[key].config(text="0")
            return

        hoje = datetime.now().date()
        # Calcula o início da semana (segunda-feira)
        inicio_semana = hoje - timedelta(days=hoje.weekday())
        inicio_mes = hoje.replace(day=1)

        total_tickets = df.shape[0]
        df_treated = df[df['status'].isin(['Resolvido', 'Fechado'])]

        treated_today = df_treated[df_treated['data'].dt.date == hoje].shape[0]
        treated_week = df_treated[df_treated['data'].dt.date >= inicio_semana].shape[0]
        treated_month = df_treated[df_treated['data'].dt.date >= inicio_mes].shape[0]

        self.stats_labels["total_tickets"].config(text=str(total_tickets))
        self.stats_labels["treated_today"].config(text=str(treated_today))
        # Formata a data para exibir no balão da semana
        self.stats_labels["treated_week"].config(text=f"{treated_week} ({inicio_semana.strftime('%d/%m')}-{ (inicio_semana + timedelta(days=6)).strftime('%d/%m')})")
        self.stats_labels["treated_month"].config(text=f"{treated_month} ({inicio_mes.strftime('%m/%Y')})")


    def _show_chart_popup(self):
        """Exibe o gráfico dinâmico e interativo em uma nova janela pop-up."""
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

        chart_window = tk.Toplevel(self.root)
        chart_window.title("Resumo e Gráficos de Tickets")
        chart_window.state("zoomed")

        # Frame para controles de filtro
        filter_controls_frame = ttk.Frame(chart_window, padding=10)
        filter_controls_frame.pack(fill="x", pady=(0, 10))

        ttk.Label(filter_controls_frame, text="Filtrar por Período:").pack(side="left", padx=5)
        # Opções de filtro para o gráfico: Total (por status), Mês, Ano
        period_options = ["Total", "Mês", "Ano"]
        self.chart_period_combobox = ttk.Combobox(filter_controls_frame, values=period_options, state="readonly")
        self.chart_period_combobox.set("Total") # Padrão: mostrar por status e total geral
        self.chart_period_combobox.pack(side="left", padx=5)

        # Configurar o gráfico inicial
        self.fig_chart = plt.figure(figsize=(10, 6))
        self.chart_ax = self.fig_chart.add_subplot(111)
        self.chart_canvas = FigureCanvasTkAgg(self.fig_chart, master=chart_window)
        self.chart_canvas_widget = self.chart_canvas.get_tk_widget()
        self.chart_canvas_widget.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=20, pady=10)

        # Atualiza o gráfico inicialmente com os filtros padrão
        self._update_chart(df, self.chart_canvas, self.chart_ax)

        # Bind para atualizar o gráfico automaticamente quando o combobox de período muda
        self.chart_period_combobox.bind("<<ComboboxSelected>>", lambda event: self._update_chart(df, self.chart_canvas, self.chart_ax))


    def _update_chart(self, df, canvas, ax):
        """Atualiza o gráfico com base nos filtros selecionados, mostrando status e total."""
        ax.clear() # Limpa o gráfico anterior

        selected_period = self.chart_period_combobox.get()
        
        filtered_df = df.copy()

        if filtered_df.empty:
            ax.text(0.5, 0.5, "Não há dados para os filtros selecionados.", transform=ax.transAxes, ha="center", va="center")
            canvas.draw()
            return

        x_labels = [] # Rótulos no eixo X (Status ou Período)
        values = []   # Valores no eixo Y (Contagem de Tickets)
        title = ""

        if selected_period == "Mês":
            # Agrupa por mês e ano
            monthly_counts = filtered_df.groupby(filtered_df['data'].dt.to_period('M')).size().sort_index()
            x_labels = [p.strftime('%m/%Y') for p in monthly_counts.index]
            values = monthly_counts.values
            title = "Tickets por Mês"
        elif selected_period == "Ano":
            # Agrupa por ano
            yearly_counts = filtered_df.groupby(filtered_df['data'].dt.year).size().sort_index()
            x_labels = [str(y) for y in yearly_counts.index]
            values = yearly_counts.values
            title = "Tickets por Ano"
        else: # selected_period == "Total" (inclui status e total geral)
            # Agrupa por status
            status_counts = filtered_df['status'].value_counts().sort_index()
            x_labels = status_counts.index.tolist()
            values = status_counts.values.tolist()

            # Adiciona a coluna de "Total Geral"
            x_labels.append("Total Geral")
            values.append(filtered_df.shape[0]) # Total de todos os tickets

            title = "Total de Tickets por Status e Geral"

        # Configurar o gráfico de barras verticais
        x_pos = range(len(x_labels))
        
        # Usar um colormap para cores mais bonitas
        colors = 'skyblue' # Cor padrão
        if len(values) > 0 and max(values) > 0:
            normalized_values = [v / max(values) for v in values]
            colors = plt.cm.viridis(normalized_values)
        else:
            colors = ['lightgray'] * len(values)
        
        ax.bar(x_pos, values, color=colors) # Usando ax.bar para gráfico vertical

        ax.set_title(title, fontsize=14)
        ax.set_xlabel("Período/Status", fontsize=12) # Eixo X é o período/status
        ax.set_ylabel("Contagem de Tickets", fontsize=12) # Eixo Y é a contagem

        ax.set_xticks(x_pos)
        ax.set_xticklabels(x_labels, rotation=45, ha='right', fontsize=10) # Rotação para rótulos longos

        # Adicionar valores nas barras para melhor leitura
        for i, v in enumerate(values):
            ax.text(i, v + 0.5, str(v), color='black', ha='center', va='bottom', fontsize=9) # Posição acima da barra

        ax.grid(True, linestyle='--', alpha=0.6, axis='y') # Grade no eixo Y
        
        plt.tight_layout()
        canvas.draw()


    def _validate_date_input(self, event=None):
        """Valida o formato da data no campo de entrada."""
        date_str = self.data_entry.get().strip()
        if date_str: # Só valida se o campo não estiver vazio
            try:
                datetime.strptime(date_str, "%d/%m/%Y")
                self.data_entry.config(foreground="black") # Data válida
            except ValueError:
                messagebox.showwarning("Data inválida", "A data deve estar no formato DD/MM/AAAA. Por favor, corrija.")
                self.data_entry.config(foreground="red") # Marca o campo como inválido
                self.data_entry.focus_set() # Volta o foco para o campo de data


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

    def _apply_status_filter(self, event=None):
        """Aplica o filtro de status na tabela."""
        selected_status = self.filter_status_combobox.get()
        all_records = self.db.fetch_all_records()
        if selected_status == "Todos":
            filtered_records = all_records
        else:
            filtered_records = [record for record in all_records if record[5] == selected_status] # record[5] é o status
        self._load_table(filtered_records)


    def _validate_inputs(self):
        """Valida os campos obrigatórios e o formato da data."""
        date_str = self.data_entry.get().strip()
        if not date_str:
            messagebox.showwarning("Campos Vazios", "O campo 'Data' é obrigatório.")
            self.data_entry.focus_set()
            return False
        try:
            datetime.strptime(date_str, "%d/%m/%Y")
        except ValueError:
            messagebox.showwarning("Data inválida", "A data deve estar no formato DD/MM/AAAA. Por favor, corrija.")
            self.data_entry.focus_set()
            return False

        if not self.numero_entry.get() or not self.descricao_entry.get():
            messagebox.showwarning("Campos Vazios", "Os campos 'Nº Ticket/Chamado' e 'Descrição' são obrigatórios.")
            return False
        return True

    def _add_record(self):
        if not self._validate_inputs():
            return
        date = self.data_entry.get() # Já validado pelo _validate_inputs
        self.db.add_record(date, self.numero_entry.get(), self.descricao_entry.get(),
                           self.acao_entry.get(), self.status_combobox.get())
        messagebox.showinfo("Sucesso", "Registro adicionado com sucesso!")
        self._load_table()
        self._clear_fields()

    def _open_edit_window(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Nenhum Selecionado", "Selecione um registro para editar.")
            return

        record_id = self.tree.item(selected_item, 'values')[0]
        record_data = self.db.fetch_record_by_ticket_number(self.tree.item(selected_item, 'values')[2]) # Busca pelo número do ticket

        if not record_data:
            messagebox.showerror("Erro", "Registro não encontrado para edição.")
            return

        edit_window = tk.Toplevel(self.root)
        edit_window.title("Editar Registro")
        edit_window.transient(self.root)
        edit_window.grab_set()
        edit_window.focus_set()

        # Posicionar a janela de edição no centro da tela
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = 400
        window_height = 300
        x_pos = int((screen_width - window_width) / 2)
        y_pos = int((screen_height - window_height) / 2)
        edit_window.geometry(f"{window_width}x{window_height}+{x_pos}+{y_pos}")

        edit_frame = ttk.LabelFrame(edit_window, text="Editar Detalhes", padding=10)
        edit_frame.pack(padx=10, pady=10, fill="both", expand=True)

        labels = ["ID:", "Data (DD/MM/AAAA):", "Nº Ticket/Chamado:", "Descrição:", "Ação Realizada:", "Status:"]
        entries = {}
        status_options = ["Aguardando Parceiro", "Cancelado", "Em Andamento", "Fechado", "Pendente de Resposta", "Resolvido"]

        for i, label_text in enumerate(labels):
            ttk.Label(edit_frame, text=label_text).grid(row=i, column=0, padx=5, pady=5, sticky="w")
            if label_text == "Status:":
                combobox = ttk.Combobox(edit_frame, values=status_options, state="readonly")
                combobox.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
                combobox.set(record_data[i])
                entries[label_text] = combobox
            else:
                entry = ttk.Entry(edit_frame)
                entry.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
                entry.insert(0, record_data[i])
                entries[label_text] = entry

        # Desabilitar edição do ID
        entries["ID:"].config(state="readonly")

        def save_edit():
            new_data = entries["Data (DD/MM/AAAA):"].get()
            new_numero = entries["Nº Ticket/Chamado:"].get()
            new_descricao = entries["Descrição:"].get()
            new_acao = entries["Ação Realizada:"].get()
            new_status = entries["Status:"].get()

            # Validação de data na janela de edição
            try:
                datetime.strptime(new_data, "%d/%m/%Y")
            except ValueError:
                messagebox.showwarning("Data inválida", "A data deve estar no formato DD/MM/AAAA. Por favor, corrija.")
                return

            if not new_numero or not new_descricao:
                messagebox.showwarning("Campos Vazios", "Os campos 'Nº Ticket/Chamado' e 'Descrição' são obrigatórios.")
                return

            self.db.update_record(record_id, new_data, new_numero, new_descricao, new_acao, new_status)
            messagebox.showinfo("Sucesso", "Registro atualizado com sucesso!")
            edit_window.destroy()
            self._load_table()

        ttk.Button(edit_frame, text="Salvar", command=save_edit).grid(row=len(labels), column=0, columnspan=2, pady=10)


    def _open_delete_window(self):
        delete_window = tk.Toplevel(self.root)
        delete_window.title("Deletar Registros")
        delete_window.transient(self.root)
        delete_window.grab_set()
        delete_window.focus_set()

        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = int(screen_width * 0.8)
        window_height = int(screen_height * 0.9)
        x_pos = int((screen_width - window_width) / 2)
        y_pos = int((screen_height - window_height) / 2)
        delete_window.geometry(f"{window_width}x{window_height}+{x_pos}+{y_pos}")

        delete_tree_frame = ttk.Frame(delete_window, padding=10)
        delete_tree_frame.pack(fill="both", expand=True)

        cols = ("ID", "Data", "Nº Ticket", "Descrição", "Ação Realizada", "Status")
        delete_tree = ttk.Treeview(delete_tree_frame, columns=cols, show='headings', selectmode='extended', height=15)
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
                    record_id = delete_tree.item(item_id, 'values')[0]
                    self.db.delete_record(record_id)
                    deleted_count += 1
                messagebox.showinfo("Sucesso", f"{deleted_count} registro(s) deletado(s) com sucesso!")
                delete_window.destroy()
                self._load_table()

        delete_button_frame = ttk.Frame(delete_window, padding=10)
        delete_button_frame.pack(fill="x")
        ttk.Button(delete_button_frame, text="Deletar Selecionados", command=perform_delete).pack(side="left", padx=5)


    def _fill_fields_on_select(self, event):
        """Preenche os campos de entrada com os dados do registro selecionado na Treeview."""
        selected_item = self.tree.focus()
        if selected_item:
            values = self.tree.item(selected_item, 'values')
            self.id_entry.config(state="normal") # Habilita para preencher
            self.id_entry.delete(0, tk.END)
            self.id_entry.insert(0, values[0])
            self.id_entry.config(state="readonly") # Desabilita novamente

            self.data_entry.delete(0, tk.END)
            self.data_entry.insert(0, values[1])

            self.numero_entry.delete(0, tk.END)
            self.numero_entry.insert(0, values[2])

            self.descricao_entry.delete(0, tk.END)
            self.descricao_entry.insert(0, values[3])

            self.acao_entry.delete(0, tk.END)
            self.acao_entry.insert(0, values[4])

            self.status_combobox.set(values[5])
        else:
            self._clear_fields()

    def _search_record(self):
        """Busca registros pelo número do ticket/chamado."""
        search_term = self.numero_entry.get()
        if not search_term:
            messagebox.showwarning("Campo Vazio", "Por favor, insira um número de ticket para buscar.")
            self._load_table() # Recarrega a tabela completa se o campo de busca estiver vazio
            return

        records = self.db.search_by_number(search_term)
        if records:
            self._load_table(records)
        else:
            messagebox.showinfo("Não Encontrado", f"Nenhum registro encontrado para o número de ticket '{search_term}'.")
            self._load_table() # Recarrega a tabela completa se nada for encontrado

    def _clear_fields(self):
        """Limpa todos os campos de entrada."""
        self.id_entry.config(state="normal")
        self.id_entry.delete(0, tk.END)
        self.id_entry.config(state="readonly")
        self.data_entry.delete(0, tk.END)
        self.data_entry.insert(0, datetime.now().strftime("%d/%m/%Y")) # Volta para a data atual
        self.data_entry.config(foreground="black") # Reseta a cor da validação
        self.numero_entry.delete(0, tk.END)
        self.descricao_entry.delete(0, tk.END)
        self.acao_entry.delete(0, tk.END)
        self.status_combobox.set("Em Andamento") # Volta para o status padrão

    def _import_data(self):
        """Importa dados de um arquivo Excel (.xlsx) ou CSV (.csv)."""
        file_path = filedialog.askopenfilename(
            filetypes=[("Arquivos Excel", "*.xlsx"), ("Arquivos CSV", "*.csv")]
        )
        if not file_path:
            return

        try:
            if file_path.endswith('.xlsx'):
                df = pd.read_excel(file_path)
            elif file_path.endswith('.csv'):
                df = pd.read_csv(file_path)
            else:
                messagebox.showerror("Formato Inválido", "Por favor, selecione um arquivo .xlsx ou .csv.")
                return

            # Mapeamento de colunas para garantir que os nomes estejam corretos
            # Garante que as colunas esperadas existam ou define valores padrão
            expected_columns = {
                'data': 'data',
                'numero_ticket': 'numero_ticket',
                'descricao': 'descricao',
                'acao_realizada': 'acao_realizada',
                'status': 'status'
            }
            # Renomeia colunas do DataFrame para corresponder ao esperado, se houver diferença
            df.rename(columns={col: expected_columns[col] for col in expected_columns if col in df.columns}, inplace=True)

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

    def _export_data(self):
        """Exporta todos os dados da tabela para um arquivo CSV ou Excel."""
        records = self.db.fetch_all_records()
        if not records:
            messagebox.showinfo("Exportar Dados", "Não há dados para exportar.")
            return

        df = pd.DataFrame(records, columns=["ID", "Data", "Nº Ticket", "Descrição", "Ação Realizada", "Status"])

        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("Arquivos CSV", "*.csv"), ("Arquivos Excel", "*.xlsx")],
            title="Salvar Dados Como"
        )
        if not file_path:
            return

        try:
            if file_path.endswith('.csv'):
                df.to_csv(file_path, index=False, encoding='utf-8-sig')
            elif file_path.endswith('.xlsx'):
                df.to_excel(file_path, index=False)
            messagebox.showinfo("Exportação Concluída", f"Dados exportados com sucesso para:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Erro de Exportação", f"Ocorreu um erro ao exportar os dados: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = TicketApp(root)
    root.mainloop()
