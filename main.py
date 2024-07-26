import openpyxl
import click
import sys
import firebase_admin
from firebase_admin import credentials, firestore

# Replace with your service account key path
cred = credentials.Certificate('cred.json')
firebase_admin.initialize_app(cred)
db = firestore.client()

_indices = [
  'IDXHIDIV20',
  'ISSI',
  'JII',
  'JII70',
  'LQ45',
]

# Function to dynamically detect and extract table data
def extract_table_data(sheet):
  table_data = []
  in_table = False
  
  for row in sheet.iter_rows(values_only=True):
    # Check if the second cell in the row is a number or the row contains valid data
    if isinstance(row[1], (int, float)) and (row[1] is not None and row[1] != 'No.'):
      in_table = True
      table_data.append(row)
    elif in_table:
      break
          
  return table_data

@click.command()
def indices():
  doc_ref = db.collection('indices')
  docs = doc_ref.stream()
  for doc in docs:
    print(f'{doc.id}')

# Command to extract data from Excel file
@click.command()
@click.option('--index', required=True, type=click.Choice(_indices, case_sensitive=False), help='The indices name')
@click.option('--file', required=True, type=click.Path(exists=True), help='Path to the Excel file to extract data from')
@click.option('--sync', is_flag=True, prompt='Sync result to Firestore?', help='Sync to Firestore')
def extract(index, file, sync):
  try:
    workbook = openpyxl.load_workbook(file)
    sheet = workbook.active
    data = extract_table_data(sheet)
    companies = []
    
    for row in data:
      companies.append(row[2])

    print(index)
    print(companies)

    # Upload data to Firestore
    if sync:
      print('Sync to Firestore...')
      store_data_in_firestore(companies, index)
  except Exception as e:
    print(f"Error: {e}")
    sys.exit(1)

def store_data_in_firestore(data, index):
  """Stores the extracted data in a Firestore collection."""
  collection_name = 'indices'
  doc_ref = db.collection(collection_name).document(f'{index}')  # Generate document ID with index
  doc_ref.update({
    'companies': data
  })

@click.group()
def cli():
  pass

cli.add_command(extract)
cli.add_command(indices)
# cli.add_command(clear_db)
# cli.add_command(parse_sync_db)

if __name__ == '__main__':
  cli()
