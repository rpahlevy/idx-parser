# IDX-PARSER

Extract indices data from IDX xlsx files and upload it into Firestore.

## Prerequisites

- Python 3 (tested with 3.11.9)
- `openpyxl` library: `pip install openpyxl`
- `click` library: `pip install click`
- `firebase-admin` library: `pip install firebase-admin`
- A Firebase project with a Firestore database set up (TODO. add schema)
- Service Account key for your Firebase project (downloaded as JSON)
- Download XLSX file from: https://www.idx.co.id/id/data-pasar/data-saham/indeks-saham then put into `xlsx` directory

## Installation

1. Install the required libraries as mentioned above. Better to use either `Pipfile` or `requirements.txt`.
2. Download your Firebase service account key as a JSON file and store it securely (not in a public repository).
3. Clone this repository: `git clone https://github.com/rpahlevy/idx-parser.git`
4. Replace `cred.json` in the `main.py` with the actual path to your JSON file.

## Usage

This script provides three functionalities:

1. List existing indices:

  ```bash
  python main.py indices
  ```

  This command will list all existing documents in the `indices` collection in your Firestore database.

2. Extract and upload data:

  ```bash
  python script.py extract \
    --index INDEX_NAME \
    --file path/to/your/excel_file.xlsx \
    --sync (optional)
  ```

  Replace the following:

  - `INDEX_NAME`: The name of the index you want to extract data for (must be one of the valid options).
  - `path/to/your/excel_file.xlsx`: The path to your Excel file containing the data table.
  - `--sync` (optional): Include this flag if you want to upload the extracted data to Firestore after extraction.

3. Link indices to other data (experimental):

  ```bash
  python main.py link-indices
  ```

  This command creates or updates the `indices` field within each company document in the `companies` collection. The `indices` field is an array containing the names of the indices to which the company belongs. This linking process facilitates efficient filtering of companies based on multiple indices.

## Supported Indices

The script currently supports extracting data for the following indices:

```
IDXHIDIV20
ISSI
JII
JII70
LQ45
```

**Note**: You can easily add support for additional indices by modifying the _indices list in the script.

## Script Location

This script assumes you're running it from the project's root directory. If the script is located elsewhere, adjust the python command accordingly.

## Notes

- The script currently assumes a specific table format in your Excel file. You might need to adjust the `extract_table_data` function if the format differs.
- The script creates/updates a document in the `indices` collection with the provided index name as the document ID and stores the extracted company list as a field named `companies`.

## Contributing

Feel free to suggest improvements or report issues through pull requests or GitHub discussions.
