# BSON to JSON to XLSX Converter

This Python script helps you process and convert `.bson.gz` and `.json.gz` files into `.json` and then into `.xlsx` format. The main objective of this tool is to:
1. Extract `.bson.gz` files, convert them to `.bson`, then to `.json` using `bsondump`.
2. Extract `.json.gz` files and save them directly as `.json`.
3. Flatten the JSON data (if nested) and convert it into `.xlsx` format for easy migration.

## Requirements

Before running the script, make sure you have the following installed:

- Python 3.x
- Required Python packages:
  - `pandas`
  - `openpyxl`
  
You can install these dependencies using pip:

```bash
pip install pandas openpyxl
```

## Setup

### Step 1: Download the MongoDB Database Tools:
1. Download MongoDB Database Tools from MongoDB official website https://www.mongodb.com/try/download/relational-migrator .
2. Extract the tools and locate bsondump.exe in the bin folder.
3. Update the CAMINHO_BSONDUMP variable in the script to point to your bsondump.exe location.

### Step 2: Directory Structure:
1. Create a folder to put all the .bson.gz and .json.gz
2. Update the PASTA_ORIGEM variable in the script to point to the source directory.
3. Place all your .bson.gz and .json.gz files in the source directory.

## Usage
1. Place your .bson.gz and .json.gz files in the source directory (specified by PASTA_ORIGEM).
2. Run the Python script:
```bash
python convert_bson_json_xlsx.py
```
3. The script will process the files and organize them into the following directories:
  - json_convertidos: This folder will store the resulting .json files.
  - bson_descompactados: This folder will store the .bson files after they are extracted from .bson.gz.
  - xlsx_convertidos: This folder will store the final .xlsx files.

## Notes
- The script uses json_normalize from pandas to flatten any nested JSON objects before converting them to .xlsx.
- Ensure the MongoDB Database Tools (bsondump) are correctly installed and accessible.
- Files will be processed and saved automatically. Make sure to check the output directories for the results.

### License
This project is licensed under the MIT License - see the LICENSE.md file for details.

### Author
Ivan Gonçalves da Silva
