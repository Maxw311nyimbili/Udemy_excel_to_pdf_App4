import pandas as pd
import glob
# glob returns all file paths that match a specific pattern, in the form of a list. It, basically, iterates through a
# folder and gets file paths with a specific pattern, then returns them as a list.

filepaths = glob.glob("Invoices/*.xlsx")
print(filepaths)

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    print(df)