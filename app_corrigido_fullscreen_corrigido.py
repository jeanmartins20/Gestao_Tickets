import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
from datetime import datetime, timedelta # Import timedelta
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
            messagebox.showinfo("Resumo", "Não há dados válidos.")
            return

        hoje = datetime.now().date()
        inicio_semana = hoje - timedelta(days=hoje.weekday())
        inicio_mes = hoje.replace(day=1)

        total_dia = df[df['data'].dt.date == hoje].shape[0]
        total_semana = df[df['data'].dt.date >= inicio_semana].shape[0]
        total_mes = df[df['data'].dt.date >= inicio_mes].shape[0]
        total_geral = df.shape[0]

        df_grouped = df.groupby(df['data'].dt.date).size().reset_index(name='tickets')

        resumo = tk.Toplevel(self.root)
        resumo.title("Resumo de Tickets")
        resumo.state("zoomed")
        resumo.configure(bg="white")

        top_frame = tk.Frame(resumo, bg="white")
        top_frame.pack(fill='x', pady=20)

        def criar_balao(texto, valor, cor):
            frame = tk.Frame(top_frame, bg=cor, bd=2, relief="groove", width=200, height=80)
            frame.pack(side="left", padx=20, expand=True, fill="both")
            tk.Label(frame, text=texto, font=("Segoe UI", 12, "bold"), bg=cor, fg="white").pack(pady=5)
            tk.Label(frame, text=str(valor), font=("Segoe UI", 20, "bold"), bg=cor, fg="white").pack()

        criar_balao("Hoje", total_dia, "#3498db")
        criar_balao("Semana", total_semana, "#27ae60")
        criar_balao("Mês", total_mes, "#f39c12")
        criar_balao("Total", total_geral, "#9b59b6")

        fig, ax = plt.subplots(figsize=(12, 5))
        ax.bar(df_grouped['data'].astype(str), df_grouped['tickets'], color="#2ecc71")
        ax.set_title("Tickets Resolvidos por Dia")
        ax.set_xlabel("Data")
        ax.set_ylabel("Quantidade")
        ax.tick_params(axis='x', rotation=45)
        ax.grid(axis='y', linestyle='--', alpha=0.7)

        chart_frame = tk.Frame(resumo, bg="white")
        chart_frame.pack(fill='both', expand=True)

        canvas = FigureCanvasTkAgg(fig, master=chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True)


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

        button_frame = ttk.Frame(self.root, padding=10)
        button_frame.grid(row=1, column=0, padx=10, sticky="ew")

        ttk.Button(button_frame, text="Adicionar", command=self._add_record).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Editar", command=self._update_record).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Deletar", command=self._delete_record).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Buscar por Nº", command=self._search_record).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Limpar Campos", command=self._clear_fields).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Estatísticas", command=self._show_statistics).pack(side="left", padx=5) # This button's command needs to be implemented or removed if not needed.
        ttk.Button(button_frame, text="Resumo e Gráfico", command=self._show_summary_and_chart).pack(side="left", padx=5)


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
        self.status_combobox.set('')
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

    # Placeholder for the _show_statistics method if it's meant to be implemented later
    def _show_statistics(self):
        messagebox.showinfo("Estatísticas", "Funcionalidade de estatísticas a ser implementada.")


if __name__ == "__main__":
    root = tk.Tk()
    app = TicketApp(root)
    root.mainloop()