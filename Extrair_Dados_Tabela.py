import sqlite3
import pandas as pd

# Conectar ao banco de dados
conn = sqlite3.connect("tickets.db")

# Verificar as tabelas
query_tabelas = "SELECT name FROM sqlite_master WHERE type='table';"
tabelas = pd.read_sql(query_tabelas, conn)
print("Tabelas encontradas:", tabelas)

# Extrair dados da tabela 'registros'
df = pd.read_sql("SELECT * FROM registros", conn)

# Exportar para Excel
df.to_csv("dados_extraidos.csv", index=False)

conn.close()
