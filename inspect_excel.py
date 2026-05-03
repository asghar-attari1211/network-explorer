import pandas as pd

def inspect(file):
    print(f"\n--- {file} ---")
    try:
        xl = pd.ExcelFile(file)
        for sn in xl.sheet_names:
            print(f"Sheet: {sn}")
            df = pd.read_excel(xl, sheet_name=sn, nrows=1)
            print(f"  Cols: {list(df.columns)}")
    except Exception as e:
        print(f"Error: {e}")

inspect('/home/asghar_attari1211/RnD/VLAN List -- AI Seekho -- Sample.xlsx')
inspect('/home/asghar_attari1211/RnD/Routes_03-May-2026.xlsx')
inspect('/home/asghar_attari1211/RnD/Dependencies_03-May-2026.xlsx')
inspect('/home/asghar_attari1211/RnD/Consolidated_VLAN_Report_27-Apr-2026.xlsx')
