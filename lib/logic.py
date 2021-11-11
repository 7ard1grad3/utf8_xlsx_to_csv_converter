import csv
import re
import sys

import pandas as pd
from rich import pretty
from rich.table import Table

from config import *


class Logic:

    def __init__(self, console):
        pretty.install()
        self.console = console
        self.valid = True
        self.errors = []
        self.valid_rows = []
        # init first steps

    def validate_file(self, file: str):
        if file is not None:
            self.console.print(f"✅ Found excel file in [bold magenta]{file}[/bold magenta]")
        else:
            self.add_error(
                f"❌ Missing file. [bold magenta]Make sure to place file in {EXCEL_FOLDER} folder "
                f"with .xlsx format[/bold magenta]", "error")
        if self.is_valid():
            self.validate_and_clean_excel(file)

    def mark_as_invalid(self):
        self.valid = False

    def is_valid(self, print_table: bool = True) -> bool:
        if print_table:
            self.show_errors()
        return self.valid

    def show_errors(self):
        if len(self.errors) > 0:
            table = Table(show_header=True, header_style="bold magenta")
            table.add_column("Message", style="dim", width=60)
            table.add_column("Type", style="red")
            for error in self.errors:
                table.add_row(error['message'], error['type'])
            self.console.print(table)

    def add_error(self, message: str, message_type: str):
        message_type = message_type.upper()
        self.errors.append({
            "message": message,
            "type": message_type,
        })
        if message_type == 'ERROR':
            self.mark_as_invalid()

    @staticmethod
    def clean(text: str) -> str:
        if str(text) == 'nan':
            text = ''
        return re.sub(REGEX_FILTER, '', str(text)).strip()

    def create_result_file(self, file: str):
        try:
            with open(f"{file.lower().replace('.xlsx', '.csv')}", 'w', newline='', encoding="utf-8") as csv_file:
                writer = csv.writer(csv_file, delimiter=DELIMITER, quotechar='"')
                writer.writerows(self.valid_rows)
            self.console.print(f"✅ csv file created ([bold magenta]{file}[/bold magenta])")
        except:
            sys.exit(f"Failed to create file in directory '{file}'")

    def validate_and_clean_excel(self, file: str):
        with self.console.status("[bold green]Working on tasks..."):
            try:
                df = pd.read_excel(file, sheet_name=0)
            except FileNotFoundError as e:
                self.console.log(f"file {e.filename} not found ", style="red on white")
                return False
            data = pd.DataFrame(df)
            # add column names
            self.valid_rows.append(list(data.columns.values))
            for row in data.values:
                valid_row = []
                for column in row:
                    valid_row.append(Logic.clean(column))
                self.valid_rows.append(valid_row)
