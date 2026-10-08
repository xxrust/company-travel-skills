"""Export only paper reimbursement sheets through WPS; never submit a print job."""
from pathlib import Path
import argparse
import win32com.client

def export(source: Path, destination: Path):
    app = win32com.client.DispatchEx("ket.Application")
    book = None
    try:
        app.Visible = False
        app.DisplayAlerts = False
        book = app.Workbooks.Open(str(source.resolve()), ReadOnly=True)
        sheets = [book.Worksheets(i) for i in range(1, book.Worksheets.Count + 1)
                  if book.Worksheets(i).Name.startswith("报销单")]
        if len(sheets) != 1:
            raise ValueError("This exporter expects one paper form; export each additional form separately.")
        sheets[0].ExportAsFixedFormat(0, str(destination.resolve()))
    finally:
        if book is not None:
            book.Close(False)
        app.Quit()

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("workbook", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    export(args.workbook, args.output)
    print(args.output.resolve())
