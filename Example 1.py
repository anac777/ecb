# Import the necessary libraries
import pandas as pd
from tabulizer import *
import tabulizer
import numpy
import re

# Define the input file
try:
    input_file
except NameError:
    # Fictional name
    input_file = r"example.xlsx"
    data = {}

wbk: ExcelWorkbook = ExcelLoader.load(input_file)

print("It was uploaded the file with the name: ", wbk.book_name())

# Choose the sheet of the Excel workbook we'll be working on
for sht_name in wbk.worksheet_names():
        print("It was red a sheet of the file with the following name:", sht_name)
        sht: ExcelWorksheet = wbk.get_worksheet_by_name("ABC")

        ts = tabulizer.TableScanner()

        row_member_settings = {"left": "B", "right": "C", "top": 3}
        col_member_settings = {"top": 2, "bottom": 2, "left": "D"}

        ts.scan_table(
            sht=sht,
            row_member_settings=row_member_settings,
            row_max_skip=3,
            col_member_settings=col_member_settings,
            col_max_skip=1,
            row_member_strip_text=True,
            col_member_strip_text=True
        )

        cards = ts.get_table(
            row_member_names=["Fluxo", "Pais"],
            col_member_names=["Metrica"],
            replace_with_null=["N/A"],
            skip_nulls=False
        )

        cards.drop(columns=["Linha", "Coluna", "Coluna (texto)"], inplace=True)

        # Multiply by 1000 and change the name accordingly
        cards.loc[cards['Metrica'] == 'Value (Thousands EUR)', 'Valor'] *= 1000

        cards.loc[cards['Metrica'] == 'Value (Thousands EUR)', 'Metrica'] = 'Value (EUR)'

        # Add period column
        workbook_name = wbk.book_name()
        last_8_characters = workbook_name[-8:]
        year = last_8_characters[:4]
        month = last_8_characters[4:6]
        formatted_period = f"{year}-{month}"

        cards['Periodo'] = formatted_period

        # Re-define index
        cards.reset_index(drop=True, inplace=True)

        # Add column "Origem"
        cards['Origem'] = 'PT'
        cards.loc[cards.index > 77, 'Origem'] = 'NotPT'

        # Rows to delete based on index
        rows_to_delete = [52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 76, 77, 78, 79, 84, 85, 86, 87]

        cards_filtered = cards.drop(index=rows_to_delete)

        # Converter to float
        cards_filtered['Valor'].fillna(0, inplace=True)
        cards_filtered['Valor'] = cards_filtered['Valor'].astype(float)

        # Conditional indexing
        mask = cards_filtered['Pais'].notnull()

        cards_filtered.loc[mask, 'Fluxo'] = "Cash withdrawals abroad - " + cards_filtered.loc[mask, 'Fluxo']

        cards_filtered.loc[mask, 'Fluxo'] += " - " + cards_filtered.loc[mask, 'Pais'].str[0]

        cards_filtered.loc[mask, 'Pais'] = cards_filtered.loc[mask, 'Pais'].str[3:]

        # Variable with the sheet's name
        cards_filtered['Utilizador'] = 'ABC'

        # Reorder variables
        cards_filtered = cards_filtered.reindex(columns=['Periodo', 'Utilizador', 'Origem', 'Fluxo', 'Pais', 'Metrica', 'Valor'])

# Round to 2 decimal places
cards_filtered['Valor'] = cards_filtered['Valor'].round(2)

data['CARDS'] = cards_filtered