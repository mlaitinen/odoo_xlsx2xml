import os
import sys
from openpyxl import load_workbook
from lxml import etree


def generate_accounts(sheet, data, record_model):

    # Get the header values / field names
    headers = {cell.column: cell.value for cell in sheet.rows.__next__() if cell.value}
    id_col = None
    for col, field_name in headers.items():
        if field_name == 'id':
            id_col = col
            headers.pop(col)
            break

    if id_col is None:
        raise Exception("The ID column is missing")

    def create_account(account_row):
        # Create a single account record.
        record = etree.Element("record", model=record_model)

        for cell in account_row:
            if cell.column in headers:
                field = etree.Element("field", name=headers[cell.column])
                field.text = str(cell.value)
                record.append(field)
            elif cell.column == id_col:
                record.attrib['id'] = cell.value
        return record

    # Iterate the rows in the worksheet
    for row in sheet.iter_rows(row_offset=1):
        data.append(create_account(row))


def main():
    # Get parameters
    if len(sys.argv) < 3:
        print("Usage: {} <xlsx file> <record model> [<output.xml>]".format(os.path.basename(__file__)))
        exit()
    elif len(sys.argv) >= 4:
        output_file = sys.argv[3]
    else:
        output_file = 'output.xml'

    src_filename = sys.argv[1]

    try:
        sheet = load_workbook(filename=src_filename, use_iterators=True).worksheets[0]
    except IOError:
        print("No such file:", src_filename)
        exit()

    # Create the XML tree
    root = etree.Element("openerp")
    data = etree.Element("data", noupdate="1")
    root.append(data)
    generate_accounts(sheet, data, sys.argv[2])

    # Write the XML content to the output file
    xml_str = etree.tostring(root, pretty_print=True)
    with open(output_file, 'wb') as f:
        f.write(xml_str)

if __name__ == "__main__":
    main()