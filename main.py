import datetime
import glob
import os
import openpyxl
from pathlib import Path

from rich.console import Console

import config
from config import EXCEL_FOLDER, WAIT_FOR_EXIT
from lib.logic import Logic


def init_directories():
    dirs = [EXCEL_FOLDER, f"{EXCEL_FOLDER}/complete", f"{EXCEL_FOLDER}/results"]
    for dir in dirs:
        if not os.path.exists(dir):
            os.makedirs(dir)



def config_file():
    path_to_text_file = "allowed_chars.txt"
    if not os.path.exists(path_to_text_file):
        with open(path_to_text_file, "w+", encoding='utf-8') as f:
            f.write(config.REGEX_FILTER)


if __name__ == '__main__':
    init_directories()
    config_file()
    console = Console()
    # Check all excel files in excel folder

    for file in glob.glob(f"{EXCEL_FOLDER}/*.xlsx"):
        file_path = f"{file}"
        app = Logic(console)
        app.validate_file(file_path)
        if app.is_valid():
            # convert to csv
            filename = file.split("/")[-1].split("\\")[-1]
            app.create_result_file(
                f"{EXCEL_FOLDER}/results/{datetime.datetime.now().strftime('%d-%m-(%H-%M-%S)')}_{filename}")
            # move file to complete files
            try:
                Path(file_path).rename(f'{EXCEL_FOLDER}/complete/{filename}')
            except FileExistsError as e:
                app.console.print(f'❌ Failed to move file {e.filename} to folder complete. File already exists')

        else:
            app.console.print('❌ Excel file is invalid and cannot be processed')
    console.print('✅ All files completed')

    if WAIT_FOR_EXIT:
        input("Press enter to close")
