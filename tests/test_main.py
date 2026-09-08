import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook
from sqlalchemy import create_engine, inspect, text

from main import import_directory, normalize_table_name


class MainTests(unittest.TestCase):
	def test_normalize_table_name_uses_lowercase_filename(self):
		self.assertEqual(normalize_table_name(Path('BASEINV.xlsx')), 'baseinv')

	def test_import_directory_creates_table_and_appends_rows(self):
		with tempfile.TemporaryDirectory() as temp_dir:
			base_path = Path(temp_dir)
			tables_dir = base_path / 'tabelas'
			tables_dir.mkdir()
			workbook_path = tables_dir / 'BASEINV.xlsx'
			database_path = base_path / 'dados.db'

			workbook = Workbook()
			worksheet = workbook.active
			worksheet.append(['ID', 'NUMINV', 'STATUS'])
			worksheet.append([1, 2617, '<'])
			workbook.save(workbook_path)
			workbook.close()

			database_url = f'sqlite:///{database_path}'

			first_import = import_directory(tables_dir, database_url)
			second_import = import_directory(tables_dir, database_url)

			self.assertEqual(first_import, {'BASEINV.xlsx': 1})
			self.assertEqual(second_import, {'BASEINV.xlsx': 1})

			engine = create_engine(database_url)
			inspector = inspect(engine)
			self.assertIn('baseinv', inspector.get_table_names())
			self.assertEqual(
				[column['name'] for column in inspector.get_columns('baseinv')],
				['ID', 'NUMINV', 'STATUS'],
			)

			with engine.connect() as connection:
				row_count = connection.execute(text('SELECT COUNT(*) FROM baseinv')).scalar_one()

			self.assertEqual(row_count, 2)


if __name__ == '__main__':
	unittest.main()
