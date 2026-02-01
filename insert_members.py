import pandas as pd
from sqlalchemy import create_engine

file_path = "members.xlsx"  
df = pd.read_excel(file_path)

def split_name(name):
    parts = str(name).split()
    first = parts[0] if len(parts) > 0 else ""
    middle = parts[1] if len(parts) > 2 else ""
    last = parts[-1] if len(parts) > 1 else ""
    return pd.Series([first, middle, last])

df[['FirstName', 'MiddleName', 'LastName']] = df['MemberName'].apply(split_name)

columns_to_insert = [
    'FirstName', 'MiddleName', 'LastName',
    'SycGenderId', 'SycMemberType', 'UsmOffice'
]

df_db = df[columns_to_insert]

engine = create_engine(
    "mysql+mysqlconnector://root:1234@localhost:3306/NexGenCoSysDBDev"
)

try:
    df_db.to_sql('MemMemberRegistration', con=engine, if_exists='append', index=False)
    print(" Data inserted successfully!")
except Exception as e:
    print(" Error:", e)
