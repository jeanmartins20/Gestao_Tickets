import tkinter as tk
from tkinter import ttk, messagebox, filedialog # Importa filedialog para seleção de arquivos
import sqlite3
from datetime import datetime, timedelta
import pandas as pd # Importa pandas para ler arquivos CSV/Excel
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


class TicketApp:
    def __init__(self, root_window):
        self.root = root_window
        self.root.title("Gestão de Tickets de Suporte")
        self.root.state("zoomed")  # Tela cheia
        self.db = DatabaseManager('tickets.db')
        self._create_widgets()
        self._load_table()

    def _show_summary_and_chart(self):
        records = self.db.fetch_all_records()
        if not records:
            messagebox.showinfo("Resumo", "Não há dados para análise.")
            return

        df = pd.DataFrame(records, columns=["id", "data", "numero", "descricao", "acao", "status"])
        df['data'] = pd.to_datetime(df['data'], format='%d/%m/%Y', errors='coerce')
        df.dropna(subset=['data'], inplace=True)

        if df.empty:
            messagebox.showinfo("Resumo", "Não há dados válidos para análise após a limpeza.")
            return

        # Calcular tickets resolvidos
        df_resolvidos = df[df['status'] == 'Concluído']

        hoje = datetime.now().date()
        inicio_semana = hoje - timedelta(days=hoje.weekday())
        inicio_mes = hoje.replace(day=1)

        total_tickets = df.shape[0]
        total_resolvidos_dia = df_resolvidos[df_resolvidos['data'].dt.date == hoje].shape[0]
        total_resolvidos_semana = df_resolvidos[df_resolvidos['data'].dt.date >= inicio_semana].shape[0]
        total_resolvidos_mes = df_resolvidos[df_resolvidos['data'].dt.date >= inicio_mes].shape[0]

        # Criar a janela de resumo
        resumo_window = tk.Toplevel(self.root)
        resumo_window.title("Resumo e Gráfico de Tickets")
        resumo_window.state("zoomed") # Abrir em tela cheia

        # Frame para os balões de resumo
        summary_frame = ttk.Frame(resumo_window, padding=10)
        summary_frame.pack(pady=10)

        # Função auxiliar para criar os balões
        def create_summary_balloon(parent, text, value, color):
            frame = ttk.Frame(parent, relief="solid", borderwidth=1, style=f"{color}.TFrame")
            frame.pack(side="left", padx=10, pady=5, fill="both", expand=True)
            ttk.Label(frame, text=text, font=('Segoe UI', 10, 'bold'), foreground="white", background=color).pack(pady=2)
            ttk.Label(frame, text=value, font=('Segoe UI', 18, 'bold'), foreground="white", background=color).pack(pady=2)

        # Estilos para os balões
        self.root.style = ttk.Style()
        self.root.style.theme_use('clam') # Um tema que permite mais customização
        self.root.style.configure('blue.TFrame', background='#3498db') # Azul
        self.root.style.configure('green.TFrame', background='#2ecc71') # Verde
        self.root.style.configure('orange.TFrame', background='#e67e22') # Laranja
        self.root.style.configure('purple.TFrame', background='#9b59b6') # Roxo

        create_summary_balloon(summary_frame, "Total Geral", total_tickets, "blue")
        create_summary_balloon(summary_frame, "Resolvidos Hoje", total_resolvidos_dia, "green")
        create_summary_balloon(summary_frame, "Resolvidos Semana", total_resolvidos_semana, "orange")
        create_summary_balloon(summary_frame, "Resolvidos Mês", total_resolvidos_mes, "purple")

        # Gerar dados para o gráfico de colunas: Tickets por Status
        status_counts = df['status'].value_counts()
        labels = status_counts.index.tolist()
        values = status_counts.values.tolist()

        # Configurar o gráfico de colunas
        fig_bar, ax_bar = plt.subplots(figsize=(10, 6))
        ax_bar.bar(labels, values, color=['steelblue', 'olivedrab', 'indianred', 'mediumorchid', 'darkorange'])
        ax_bar.set_title("Tickets por Status")
        ax_bar.set_xlabel("Status")
        ax_bar.set_ylabel("Número de Tickets")
        ax_bar.grid(axis='y', linestyle='--', alpha=0.7)
        plt.xticks(rotation=45, ha='right') # Rotacionar rótulos para melhor visualização
        plt.tight_layout()

        # Integrar o gráfico no Tkinter
        canvas_bar = FigureCanvasTkAgg(fig_bar, master=resumo_window)
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

    def _create_widgets(self):
        input_frame = ttk.LabelFrame(self.root, text="Detalhes do Ticket", padding=15)
        input_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        input_frame.columnconfigure(1, weight=1)
        input_frame.columnconfigure(3, weight=1)
        input_frame.columnconfigure(5, weight=1)

        ttk.Label(input_frame, text="ID:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.id_entry = ttk.Entry(input_frame, width=10, state="readonly")
        self.id_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(input_frame, text="Data (DD/MM/AAAA):").grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.data_entry = ttk.Entry(input_frame)
        self.data_entry.grid(row=0, column=3, padx=5, pady=5, sticky="ew")
        # Preencher a data automaticamente com a data atual
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
        status_options = ["Em Andamento", "Pendente", "Concluído", "Cancelado", "Aguardando Parceiro"]
        self.status_combobox = ttk.Combobox(input_frame, values=status_options, state="readonly")
        self.status_combobox.grid(row=2, column=5, padx=5, pady=5, sticky="ew")
        self.status_combobox.set("Em Andamento") # Define um status inicial

        button_frame = ttk.Frame(self.root, padding=10)
        button_frame.grid(row=1, column=0, padx=10, sticky="ew")

        ttk.Button(button_frame, text="Adicionar", command=self._add_record).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Editar", command=self._update_record).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Deletar", command=self._delete_record).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Buscar por Nº", command=self._search_record).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Limpar Campos", command=self._clear_fields).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Resumo e Gráfico", command=self._show_summary_and_chart).pack(side="left", padx=5)
        # NOVO: Botão para importar dados
        ttk.Button(button_frame, text="Importar Dados", command=self._import_data).pack(side="left", padx=5)


        tree_frame = ttk.Frame(self.root)
        tree_frame.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")
        self.root.rowconfigure(2, weight=1)
        self.root.columnconfigure(0, weight=1)

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

    def _load_table(self, records=None):
        for row in self.tree.get_children():
            self.tree.delete(row)
        if records is None:
            records = self.db.fetch_all_records()
        if records:
            for row in records:
                self.tree.insert("", tk.END, values=row)

    def _validate_inputs(self):
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

    def _update_record(self):
        record_id = self.id_entry.get()
        if not record_id:
            messagebox.showwarning("Nenhum Registro", "Selecione um registro para editar.")
            return
        if not self._validate_inputs():
            return
        date = self._get_valid_date()
        if not date:
            return
        self.db.update_record(record_id, date, self.numero_entry.get(), self.descricao_entry.get(),
                              self.acao_entry.get(), self.status_combobox.get())
        messagebox.showinfo("Sucesso", "Registro atualizado com sucesso!")
        self._load_table()
        self._clear_fields()

    def _delete_record(self):
        record_id = self.id_entry.get()
        if not record_id:
            messagebox.showwarning("Nenhum Registro", "Selecione um registro para deletar.")
            return
        if messagebox.askyesno("Confirmar", "Tem certeza que deseja deletar este registro?"):
            self.db.delete_record(record_id)
            messagebox.showinfo("Sucesso", "Registro deletado com sucesso!")
            self._load_table()
            self._clear_fields()

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
        self.status_combobox.set('Em Andamento') # Resetar para o status padrão
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
                    # Você pode optar por mostrar um messagebox aqui, mas para muitas linhas, um log é melhor
                    # messagebox.showwarning("Erro de Linha", f"Falha ao importar linha {index+2}: {e}")

            messagebox.showinfo("Importação Concluída", f"{imported_count} registros importados com sucesso!")
            self._load_table() # Recarrega a tabela para mostrar os novos dados

        except Exception as e:
            messagebox.showerror("Erro de Importação", f"Ocorreu um erro ao importar o arquivo: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = TicketApp(root)
    root.mainloop()

