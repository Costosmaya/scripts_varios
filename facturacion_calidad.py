from tracemalloc import start
import pandas as pd
import pygsheets as gc

def main():
    json_auth_path = "C://Users//User//Documents//Analisis  Desarrollo Costos//Scripts//Python//secure_path//gcpApikey.json"
    
    gcon = gc.authorize(service_file = json_auth_path)

    wsheet = gcon.open_by_url('https://docs.google.com/spreadsheets/d/1dmZBHHVc4sRIP7ytBL_8v1QA-cu1fkh4Gm43l5p169k/edit#gid=1867979157')

    sheet = wsheet.worksheet_by_title('Facturación')

    path = 'C://Users//User//Desktop//Facturacion  Consolidado.xlsx'

    df_facturacion = pd.read_excel(path)

    df_facturacion['Facturado'] = df_facturacion['Facturado'].str.replace('Q','').str.replace(',','.').astype(float)

    df_facturacion = df_facturacion[['Fecha','Facturado']].groupby('Fecha').sum()

    sheet.clear(start='A2', end =f'B{sheet.rows}')

    sheet.set_dataframe(df_facturacion, (2,1), copy_index=True, copy_head=False)



if __name__ == '__main__':
    main()

