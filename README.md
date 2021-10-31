# XLSX to CSV converter

Convert xlsx files to csv with UTF-8 encoding and allows only hebrew and english chars 

## Installation
**place xlsx files in Excel folder.**
###via binary
* run **converter.exe** or **converter_silent.exe** (if you want to close the console right after conversion )

###via python
Install all the required packages from the requirements.txt

```bash
pip install -r requirements.txt
python3 main.py
```


**csv files will be stored at /excel/results directory**

**converted xlsx files will be moved to /excel/complete folder**

## Configurations

Can be changed in the config.py file

```python
# directory for excel files
EXCEL_FOLDER = 'excel'
# delimiter
DELIMITER = ','
# allowed chard regex
REGEX_FILTER = '[^/\\ A-Z,a-z0-9 א-ת]+'
```
