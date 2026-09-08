import tempfile
import unittest
from pathlib import Path
import sys

from openpyxl import Workbook
from sqlalchemy import create_engine, inspect, text

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from main import import_directory, normalize_table_name


class MainTests(unittest.TestCase):
	def test_normalize_table_name_uses_lowercase_filename(self):
		self.assertEqual(normalize_table_name(Path('BASEINV.xlsx')), 'baseinv')

	def test_import_directory_creates_table_with_id_pk_and_preserves_sheet_id(self):
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
			pk = inspector.get_pk_constraint('baseinv')
			self.assertEqual(pk.get('constrained_columns'), ['id'])
			self.assertEqual(
				[column['name'] for column in inspector.get_columns('baseinv')],
				['id', 'source_id', 'numinv', 'status'],
			)

			with engine.connect() as connection:
				row_count = connection.execute(text('SELECT COUNT(*) FROM baseinv')).scalar_one()
				rows = connection.execute(
					text('SELECT id, source_id, numinv, status FROM baseinv ORDER BY id')
				).fetchall()

			self.assertEqual(row_count, 2)
			self.assertEqual(rows[0][1:], (1, 2617, '<'))
			self.assertEqual(rows[1][1:], (1, 2617, '<'))

	def test_import_directory_migrates_existing_table_without_pk(self):
		with tempfile.TemporaryDirectory() as temp_dir:
			base_path = Path(temp_dir)
			tables_dir = base_path / 'tabelas'
			tables_dir.mkdir()
			workbook_path = tables_dir / 'BASEINV.xlsx'
			database_path = base_path / 'dados.db'

			workbook = Workbook()
			worksheet = workbook.active
			worksheet.append(['NUMINV', 'STATUS'])
			worksheet.append([2617, '<'])
			workbook.save(workbook_path)
			workbook.close()

			database_url = f'sqlite:///{database_path}'
			engine = create_engine(database_url)

			with engine.begin() as connection:
				connection.execute(text('CREATE TABLE baseinv (numinv INTEGER)'))
				connection.execute(text('INSERT INTO baseinv (numinv) VALUES (999)'))

			result = import_directory(tables_dir, database_url)
			self.assertEqual(result, {'BASEINV.xlsx': 1})

			inspector = inspect(engine)
			pk = inspector.get_pk_constraint('baseinv')
			self.assertEqual(pk.get('constrained_columns'), ['id'])
			self.assertEqual(
				[column['name'] for column in inspector.get_columns('baseinv')],
				['id', 'numinv', 'status'],
			)

			with engine.connect() as connection:
				rows = connection.execute(
					text('SELECT id, numinv, status FROM baseinv ORDER BY id')
				).fetchall()

			self.assertEqual(len(rows), 2)
			self.assertEqual(rows[0][1:], (999, None))
			self.assertEqual(rows[1][1:], (2617, '<'))


if __name__ == '__main__':
	unittest.main()
