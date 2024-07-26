import openpyxl
import argparse
import sys

# Load the workbook and select the active sheet
# file_path = 'xlsx/IDX HIDIV20 - Jan 2024.xlsx'
# workbook = openpyxl.load_workbook(file_path)
# sheet = workbook.active

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

# Main function to handle the script logic
def main():
  parser = argparse.ArgumentParser(description='Extract table data from an Excel file.')
  parser.add_argument('file_path', nargs='?', help='Path to the Excel file')
  args = parser.parse_args()
  
  if not args.file_path:
    print('Warning: No file path specified.')
    sys.exit(1)
  
  # Load the workbook and select the active sheet
  workbook = openpyxl.load_workbook(args.file_path)
  sheet = workbook.active
  
  # Extract the data
  data = extract_table_data(sheet)
  
  # Print the data
  for row in data:
    print(row)

if __name__ == '__main__':
  main()
