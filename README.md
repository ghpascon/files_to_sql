# files_to_sql

Projeto em Poetry para importar arquivos `.xlsx` da pasta `tabelas/` para um banco SQL usando SQLAlchemy.

## Requisitos

- Python 3.12+
- Poetry

## Instalação

```bash
poetry install
```

## URLs de banco suportadas pelo SQLAlchemy

Exemplos:

- SQLite: `sqlite:///files_to_sql.db`
- PostgreSQL: `postgresql+psycopg://usuario@localhost:5432/meu_banco`
- MySQL: `mysql+pymysql://usuario@localhost:3306/meu_banco`
- SQL Server: `mssql+pyodbc://usuario@localhost/meu_banco?driver=ODBC+Driver+18+for+SQL+Server`

Você também pode passar a conexão por variável de ambiente:

```bash
export DATABASE_URL="sqlite:///files_to_sql.db"
```

## Uso

Coloque os arquivos `.xlsx` (ou `.xlxs`, se vierem com essa extensão) dentro da pasta `tabelas/` e execute:

```bash
poetry run python main.py
```

Ou informe os parâmetros diretamente:

```bash
poetry run python main.py --database-url "sqlite:///files_to_sql.db" --tables-dir "tabelas"
```

## Comportamento

- Cada arquivo `.xlsx` vira uma tabela com o nome do arquivo em minúsculas
- Exemplo: `BASEINV.xlsx` gera a tabela `baseinv`
- A primeira linha da planilha vira o cabeçalho das colunas
- Os dados são inseridos sem apagar a tabela existente
- Se a tabela já existir e faltarem colunas novas da planilha, elas são adicionadas