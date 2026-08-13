import os
import json
import pandas as pd

def converter_excel_para_json():
    # Procurar arquivo excel na pasta atual
    arquivos = [f for f in os.listdir('.') if f.endswith('.xlsx') or f.endswith('.xls')]
    if not arquivos:
        print("Erro: Nenhum arquivo .xlsx ou .xls encontrado na pasta!")
        return

    excel_file = None
    for f in arquivos:
        if 'BASE' in f.upper() or 'SAUDE' in f.upper() or 'SAÚDE' in f.upper():
            excel_file = f
            break
    if not excel_file:
        excel_file = arquivos[0]

    print(f"Lendo planilha: {excel_file}...")
    xls = pd.ExcelFile(excel_file)
    
    target_sheet = None
    for sheet in xls.sheet_names:
        if 'SAÚDE' in sheet.upper() or 'SAUDE' in sheet.upper():
            target_sheet = sheet
            break
    if not target_sheet:
        target_sheet = xls.sheet_names[0]

    print(f"Utilizando a aba: '{target_sheet}'")
    df = pd.read_excel(excel_file, sheet_name=target_sheet)

    # Remover colunas vazias
    df = df.dropna(how='all', axis=1)
    df = df.loc[:, ~df.columns.astype(str).str.startswith('Unnamed')]

    registros = []
    for idx, row in df.iterrows():
        item = {}
        for col in df.columns:
            val = row[col]
            col_name = str(col).strip()
            
            if pd.isna(val):
                item[col_name] = ""
            elif isinstance(val, pd.Timestamp):
                item[col_name] = val.strftime('%Y-%m-%d')
            elif isinstance(val, float):
                if val.is_integer():
                    item[col_name] = int(val)
                else:
                    item[col_name] = round(val, 4)
            else:
                item[col_name] = str(val).strip()
        
        if 'CNPJ' in item and item['CNPJ']:
            registros.append(item)

    out_file = 'dados.json'
    with open(out_file, 'w', encoding='utf-8') as f:
        json.dump(registros, f, ensure_ascii=False, indent=2)

    tamanho_mb = os.path.getsize(out_file) / (1024 * 1024)
    print(f"Sucesso! {len(registros)} cotações exportadas para '{out_file}' ({tamanho_mb:.2f} MB).")

if __name__ == '__main__':
    converter_excel_para_json()
