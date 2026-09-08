from __future__ import annotations

import re
from collections.abc import Iterable
from datetime import date, datetime, time
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from sqlalchemy import Column, MetaData, Table, Text, create_engine, inspect, text

DATABASE_URL_EXAMPLES = {
	'SQLite': 'sqlite:///files_to_sql.db',
	'PostgreSQL': 'postgresql+psycopg://usuario@localhost:5432/meu_banco',
	'MySQL': 'mysql+pymysql://usuario@localhost:3306/meu_banco',
	'SQL Server': 'mssql+pyodbc://usuario@localhost/meu_banco?driver=ODBC+Driver+18+for+SQL+Server',
}

DATABASE_URL = 'mysql+pymysql://root:admin@localhost:3306/idl'
# Configure os valores aqui no topo em vez de usar argumentos de linha de comando
TABLES_DIR = Path('tabelas')


def normalize_table_name(file_path: Path) -> str:
	table_name = re.sub(r'\W+', '_', file_path.stem.strip().lower()).strip('_')
	if not table_name:
		raise ValueError(f'Não foi possível gerar um nome de tabela para {file_path.name}.')
	return table_name


def normalize_column_names(headers: Iterable[Any]) -> list[str]:
	normalized: list[str] = []
	seen: dict[str, int] = {}

	for index, header in enumerate(headers, start=1):
		value = str(header).strip() if header not in (None, '') else f'coluna_{index}'
		if not value:
			value = f'coluna_{index}'

		count = seen.get(value, 0)
		seen[value] = count + 1
		normalized.append(value if count == 0 else f'{value}_{count + 1}')

	return normalized


def normalize_cell_value(value: Any) -> Any:
	if isinstance(value, datetime):
		return value.isoformat(sep=' ', timespec='seconds')
	if isinstance(value, date):
		return value.isoformat()
	if isinstance(value, time):
		return value.isoformat(timespec='seconds')
	return value


def create_engine_from_url(database_url: str):
	return create_engine(database_url)


def ensure_table(engine, table_name: str, columns: list[str]) -> Table:
	inspector = inspect(engine)
	metadata = MetaData()

	if table_name not in inspector.get_table_names():
		table = Table(table_name, metadata, *(Column(column, Text) for column in columns))
		metadata.create_all(engine, tables=[table], checkfirst=True)
		return table

	existing_columns = {column['name'] for column in inspector.get_columns(table_name)}
	missing_columns = [column for column in columns if column not in existing_columns]

	if missing_columns:
		preparer = engine.dialect.identifier_preparer
		quoted_table = preparer.quote(table_name)
		with engine.begin() as connection:
			for column in missing_columns:
				quoted_column = preparer.quote(column)
				connection.execute(
					text(f'ALTER TABLE {quoted_table} ADD COLUMN {quoted_column} TEXT')
				)

	metadata.reflect(bind=engine, only=[table_name])
	return metadata.tables[table_name]


def import_workbook(engine, file_path: Path) -> int:
	workbook = load_workbook(filename=file_path, read_only=True, data_only=True)
	try:
		worksheet = workbook.active
		rows = worksheet.iter_rows(values_only=True)
		headers_row = next(rows, None)

		if headers_row is None:
			return 0

		columns = normalize_column_names(headers_row)
		table = ensure_table(engine, normalize_table_name(file_path), columns)
		payload = [
			{
				column: normalize_cell_value(value)
				for column, value in zip(columns, row, strict=False)
			}
			for row in rows
			if any(value is not None and value != '' for value in row)
		]

		with engine.begin() as connection:
			if payload:
				connection.execute(table.insert(), payload)

		return len(payload)
	finally:
		workbook.close()


def import_directory(directory: Path, database_url: str) -> dict[str, int]:
	engine = create_engine_from_url(database_url)
	imported_rows: dict[str, int] = {}

	for file_path in sorted(directory.iterdir()):
		if file_path.suffix.lower() not in {'.xlsx', '.xlxs'} or not file_path.is_file():
			continue
		imported_rows[file_path.name] = import_workbook(engine, file_path)

	return imported_rows


# Argument parsing removed — configure `DATABASE_URL` and `TABLES_DIR` at the top.


def main() -> int:
	tables_dir = TABLES_DIR if isinstance(TABLES_DIR, Path) else Path(TABLES_DIR)

	if not tables_dir.exists():
		raise FileNotFoundError(f'Pasta não encontrada: {tables_dir}')

	imported_rows = import_directory(tables_dir, DATABASE_URL)

	if not imported_rows:
		print('Nenhum arquivo .xlsx foi encontrado para importar.')
		return 0

	for filename, row_count in imported_rows.items():
		print(f'{filename}: {row_count} linhas importadas.')
	return 0


if __name__ == '__main__':
	raise SystemExit(main())
