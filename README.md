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
3. Replace 'cred.json' in the script with the actual path to your JSON file.

## Usage

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

## Notes

- The script currently assumes a specific table format in your Excel file. You might need to adjust the `extract_table_data` function if the format differs.
- The script creates/updates a document in the `indices` collection with the provided index name as the document ID and stores the extracted company list as a field named `companies`.

## Contributing

Feel free to suggest improvements or report issues through pull requests or GitHub discussions.
