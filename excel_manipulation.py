import openpyxl
import pandas as pd
import os
 
def modify_excel(df_with_dispo, df_without_dispo):
 
    filename = "C:/Users/B0640469/Downloads/temp.xlsx"
 
    # write dataframe with disposition to excel worksheet
    df_with_dispo[0].to_excel(filename)
    #os.startfile("C:/Users/B0640469/Downloads/temp.xlsx")
 
    wb = openpyxl.load_workbook(filename)
    ws = wb.worksheets[0]
 
    NCRs_with_dispo = df_with_dispo[0]["NCR Number"].tolist()
    for index, row in df_without_dispo[0].iterrows():
        if row["NCR Number"] not in NCRs_with_dispo:
            print(row["NCR Number"])
            modified_row = row.tolist()
            modified_row.insert(0,0)
            modified_row.insert(14, "PENDING DISPO")
            ws.append(modified_row)
    wb.save(filename)