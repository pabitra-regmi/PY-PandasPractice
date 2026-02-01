import os
import pandas as pd
from tkinter import Tk, filedialog

root = Tk()
root.withdraw()

file_path = filedialog.askopenfilename(
    title="Select a data file",
    filetypes=[
        ("CSV", "*.csv"),
        ("Excel", "*.xlsx;*.xls"),
        ("JSON", "*.json"),
        ("Text", "*.txt"),
        ("All Files", "*.*")
    ]
)
root.destroy()

if not file_path:
    print("No file selected.")
    exit()

print("Selected:", file_path)

ext = os.path.splitext(file_path)[1].lower()

if ext in [".xlsx", ".xls"]:
    df = pd.read_excel(file_path)

elif ext == ".csv":
    df = pd.read_csv(file_path)

elif ext == ".json":
    df = pd.read_json(file_path)

elif ext == ".txt":
    df = pd.read_csv(file_path, sep=None, engine="python")  

else:
    print("Unsupported file type for pandas.")
    exit()

print("\n File loaded successfully using pandas!")
print(df.head())

