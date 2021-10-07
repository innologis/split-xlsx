import openpyexcel
import argparse
import sys

def not_empty_row(row):
    for cell in row:
        if cell.value is not None:
            return True


def parse_xlsx(filename):
    inf = open(filename, "rb") if filename else sys.stdin
    wb = openpyexcel.load_workbook(inf, read_only=True)
    ws = iter(wb[wb.sheetnames[0]])
    headers = [cell.value for cell in next(ws)]
    return headers, [
        dict(zip(headers, [cell.value for cell in row]))
        for row in ws
        if not_empty_row(row)
    ]


def store_xlsx(filename, headers, data):
    wb = openpyexcel.Workbook()
    ws = wb.active
    ws.append(headers)
    for row in data:
        row_data = []
        for h in headers:
            row_data.append(row[h] or "")
        ws.append(row_data)
    wb.save(filename)


def uniq_boxes(data, field):
    fr = list(map(lambda x: x[field], data))
    return set(fr)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("-i", "--input", help="Input file", required=True)
    parser.add_argument("-f", "--fieldname", help="Field name", required=True)
    parser.add_argument("-o", "--outdir", help="Path to folder to place results in")

    args = parser.parse_args()

    headers, data = parse_xlsx(args.input)
    h = list(filter(lambda x: x is not None, headers))

    boxes = uniq_boxes(data, args.fieldname)
    od = args.outdir + "/" if args.outdir else "./"
    
    for box in boxes:
        box_content = list(filter(lambda x: x[args.fieldname] == box, data))
        store_xlsx(od + box + ".xlsx", h, box_content)


if __name__ == "__main__":
    main()
