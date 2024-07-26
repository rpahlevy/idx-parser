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
_doc_indices = 'indices'
_doc_companies = 'companies'
_field_indices = 'indices'
_field_companies = 'companies'

@click.command()
def indices():
  docs = get_indices()  
  for doc in docs:
    print(f'{doc.id}')

def get_indices():
  doc_ref = db.collection(_doc_indices)
  docs = doc_ref.stream()
  return docs

def get_companies():
  doc_ref = db.collection(_doc_companies)
  docs = doc_ref.stream()
  return docs

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
      store_company_in_index(companies, index)
  except Exception as e:
    print(f"Error: {e}")
    sys.exit(1)

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

def store_company_in_index(data, index):
  """Stores the extracted data in a Firestore collection."""
  doc_ref = db.collection(_doc_indices).document(f'{index}')  # Generate document ID with index
  doc_ref.update({
    _field_companies: data
  })

# Command to link indices to companies
@click.command()
@click.option('--clear', is_flag=True, prompt='Clear previous indices list in each company?')
def link_indices(clear):
  if (clear):
    clear_company_indices()

  # get indices collection
  indices = get_indices()
  companies = {}
  for i in indices:
    id = i.id
    index = i.to_dict()

    # filter out indices without companies
    if _field_companies not in index:
      continue

    # foreach company in the index,
    # add it to companies dict and append the index to it
    for company in index[_field_companies]:
      if company not in companies:
        companies[company] = {
          _field_indices: []
        }
      companies[company][_field_indices].append(id)

  docs = db.collection(_doc_companies).stream()
  batch = db.batch()

  for doc in docs:
    if doc.id not in companies:
      continue

    doc_ref = doc.reference
    batch.update(doc_ref, companies[doc.id])

  batch.commit()

def clear_company_indices():
  """Clears the specified field for all documents in a collection using batch writes."""

  docs = db.collection(_doc_companies).stream()
  batch = db.batch()

  for doc in docs:
    doc_ref = doc.reference
    batch.update(doc_ref, {_field_indices: []})

  batch.commit()


@click.group()
def cli():
  pass

cli.add_command(extract)
cli.add_command(indices)
cli.add_command(link_indices)
# cli.add_command(clear_db)
# cli.add_command(parse_sync_db)

if __name__ == '__main__':
  cli()
