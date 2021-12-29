import camelot
import sys

def extract_pdf(filepath):
    tables = camelot.read_pdf(filepath)
    # filename = filepath
    tables[0].df
    tables.export('output.excel', f='excel', compress=True)
    tables[0].to_excel('output.excel')
    
if __name__ == '__main__':
    file_path = sys.argv[1]
    extract_pdf(file_path)