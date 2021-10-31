import datetime
import glob
import os
from pathlib import Path

from rich.console import Console

from config import EXCEL_FOLDER, WAIT_FOR_EXIT
from lib.logic import Logic

if __name__ == '__main__':
    console = Console()
    # Check all excel files in excel folder
    try:
        os.chdir(f"./{EXCEL_FOLDER}")
        for file in glob.glob("*.xlsx"):
            file_path = f"../excel/{file}"
            app = Logic(console)
            app.validate_file(file_path)
            if app.is_valid():
                # convert to csv
                app.create_result_file(
                    f"../{EXCEL_FOLDER}/results/{datetime.datetime.now().strftime('%d-%m-(%H-%M-%S)')}_{file}")
                # move file to complete files
                try:
                    Path(file_path).rename(f"../{EXCEL_FOLDER}/complete/{file}")
                except FileExistsError as e:
                    app.console.print(f'❌ Failed to move file {e.filename} to folder complete. File already exists')

            else:
                app.console.print('❌ Excel file is invalid and cannot be processed')
        console.print('✅ All files completed')
    except FileNotFoundError:
        console.print("No files found in excel folder")

    if WAIT_FOR_EXIT:
        input("Press enter to close")

