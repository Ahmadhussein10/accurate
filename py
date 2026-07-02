import pandas as pd
xls = pd.ExcelFile('MV_FO_Cable_Handover_Master.xlsx')

def get_clean_html(df, title):
    # drop empty rows/cols
    df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
    if df.empty:
        return f"<h2>{title}</h2><p>No Data</p>"
    
    # fix headers if needed
    if 'Unnamed' in str(df.columns[0]):
        df.columns = df.iloc[0].fillna('')
        df = df[1:]
        
    df = df.fillna('')
    
    html = f"<h2>{title}</h2>\n<table border='1' style='border-collapse:collapse;margin-bottom:20px;'>\n"
    html += "  <tr>" + "".join(f"<th style='background:#007BFF;color:white;padding:5px;'>{c}</th>" for c in df.columns) + "</tr>\n"
    
    for _, row in df.iterrows():
        html += "  <tr>"
        for val in row:
            val_str = str(val)
            if val_str.endswith('.0'): val_str = val_str[:-2]
            html += f"<td style='padding:5px;'>{val_str}</td>"
        html += "</tr>\n"
    html += "</table>\n"
    return html

for sheet in xls.sheet_names:
    df = pd.read_excel(xls, sheet_name=sheet)
    sheet_html = get_clean_html(df, sheet)
    print(f"---START_SHEET_{sheet}---")
    print(sheet_html)
    print(f"---END_SHEET_{sheet}---")
