import PyPDF2
from pathlib import Path
import csv
from collections import defaultdict
from tkinter import filedialog, Tk
import datetime

now = datetime.datetime.now()
timestamp = now.strftime("%Y%m%d")

root = Tk()
root.withdraw()

FILES = Path(filedialog.askdirectory(initialdir=Path.cwd(),
                                     title="Select folder containing PDFs"))
#FILES = Path(r"P:\Projects\200954\testing")
                                     
COLUMNS = "Filename", "Annotation Count",

results = list()
for f in FILES.rglob("*.pdf"):
    print(f"counting comments in {f.name}...")
    input1 = PyPDF2.PdfFileReader(open(f, "rb"))

    TotalCount = 0
    for page in input1.pages:
        count = 0
        try :
            for annot in page['/Annots']:
                annot_obj = annot.getObject()
                if "/AP" in annot_obj.keys():
                    count += 1
            TotalCount += count
        except:
            pass
    results.append((f.name, TotalCount))

INITIALFILE = "Annotation_Count_Result_" + timestamp + ".csv"
RESULT = filedialog.asksaveasfilename(initialdir=Path.cwd(),
                                      initialfile=INITIALFILE,
                                      title="Save result as CSV",
                                      filetypes=(("CSV File", "*.csv"),))

with open(RESULT, mode="w", encoding="utf-16le", newline="\n") as w:
    csvwriter = csv.writer(w, delimiter=",", quotechar="\"", quoting=csv.QUOTE_MINIMAL)
    csvwriter.writerow(COLUMNS)
    for row in results:
        csvwriter.writerow(row)