import time
import win32com.client as win32
from numpy import int64
import numpy
from pathlib import Path
from pywintypes import com_error
import sys
from sqlalchemy import create_engine, text
import pandas as pd
import pygsheets as gc
import datetime

win32c = win32.constants
def main():
	db_connection_str = 'mysql+pymysql://reports:cognos@192.168.1.238/mayaprin?charset=utf8'

	db_connection = create_engine(db_connection_str)

	query_str = """SELECT
	CASE
		WHEN job200.j_status = 'C' then 'Terminado'
		ELSE 'No Terminado' END AS 'Estado Pedido',
		job200.j_number AS OP, job200.j_customer AS CLIENTE, 
		SUM(charge.cg_value) AS 'SIN IVA', ROUND(SUM(charge.cg_value)*1.12,4) AS TOTAL,
		SUM(charge.cg_quantity) AS CANTIDAD, charge.cg_text AS MATERIAL,
		job200.j_rep AS VENDEDOR, job200.j_orig AS COTIZADOR,
		job200.j_booked_in AS 'FECHA CREACION'
	FROM job200
	INNER JOIN charge ON job200.j_number = charge.cg_job
	WHERE cg_status NOT IN (2,3)
	AND job200.j_status <> 'X'
	GROUP BY job200.j_number, charge.cg_text
	ORDER BY job200.j_number ASC;"""

	df_pendientes = pd.read_sql_query(text(query_str), con=db_connection)

	json_auth_path = "C://Users//User//Documents//Analisis  Desarrollo Costos//Scripts//Python//secure_path//gcpApikey.json"

	gcon = gc.authorize(service_file = json_auth_path)

	sheet = gcon.open_by_url('https://docs.google.com/spreadsheets/d/1YN6FiDWjmd1Rbsp3qzPkkCyX1br7DrsIRRpBOPejsaw')
	df_wsheet = sheet.worksheet_by_title('Consolidado Entregas Diarias').get_as_df(include_tailing_empty=False)

	df_pendientes['FECHA CREACION'] = pd.to_datetime(df_pendientes['FECHA CREACION'], errors='coerce')

	df_wsheet['Fecha Reporte'] = pd.to_datetime(df_wsheet['Fecha Reporte'], errors='coerce')

	

	def get_FechaReporte(row):
		OP = row['OP']
		df_filtered = df_wsheet[df_wsheet.OP == OP]
		
		if(df_filtered.empty):
			return pd.NaT
		else:
			df_filtered.sort_values(by='Fecha Reporte', ascending=False)
			return df_filtered.iloc[0,2]

	df_pendientes['FECHA REPORTE'] = df_pendientes.apply(get_FechaReporte, axis=1)

	df_pendientes['CANTIDAD'] = df_pendientes['CANTIDAD'].astype(int64)

	current_time = datetime.datetime.now()

	df_pendientes['Días desde Ingreso'] = current_time - df_pendientes['FECHA CREACION']

	df_pendientes['Días desde Reporte'] = current_time - df_pendientes['FECHA REPORTE']

	df_pendientes['Días desde Ingreso'] = df_pendientes['Días desde Ingreso'].astype('timedelta64[D]').fillna('')

	df_pendientes['Días desde Reporte'] = df_pendientes['Días desde Reporte'].astype('timedelta64[D]').fillna('')

	df_pendientes['FECHA REPORTE'] = df_pendientes['FECHA REPORTE'].astype(str)

	df_pendientes['FECHA REPORTE'] = df_pendientes['FECHA REPORTE'].apply(lambda x: '' if x == 'NaT' else x)

	df_pendientes['FECHA CREACION'] = df_pendientes['FECHA CREACION'].astype(str)

	path = "c://Users//User//Documents//Prueba Pendientes.xlsx"

	def get_col_widths(dataframe):
		idx_max = max([len(str(s)) for s in dataframe.index.values] + [len(str(dataframe.index.name))])

		return [idx_max] + [max([len(str(s)) for s in dataframe[col].values] + [len(col)]) for col in dataframe.columns]

	with pd.ExcelWriter(path, engine='xlsxwriter') as writer:
		workbook = writer.book
		worksheet =workbook.add_worksheet('Detalle')

		currency_format = workbook.add_format()
		currency_format.set_num_format('#,##0.00_ ;-#,##0.00')

		text_format = workbook.add_format()
		text_format.set_text_wrap()

		(max_row, max_col) = df_pendientes.shape

		

		format_settings = [{'header':column,'header_format':text_format, 'format': workbook.add_format() if column not in ['SIN IVA','TOTAL','Días desde Ingreso','Días desde Reporte'] else currency_format} for i,column in enumerate(df_pendientes.columns)]

		worksheet.add_table(1,1,max_row, max_col,{'data':df_pendientes.values.tolist(),'style':'Table Style Medium 2','columns':format_settings})

		worksheet.set_column(0, max_col-1, 12)
	
	run_excel(path, 'Detalle',['VENDEDOR'],['Estado Pedido','CLIENTE','OP','Días desde Ingreso','Días desde Reporte'], [['SIN IVA','Total SIN IVA', 'Sum','#,##0.00_ ;-#,##0.00'],
	['TOTAL','SUMA TOTAL', 'Sum', '#,##0.00_ ;-#,##0.00']],'Resumen','Resumen',2)

	time.sleep(20)

	run_excel(path, 'Detalle',[],['VENDEDOR','Estado Pedido'], [['SIN IVA','Total SIN IVA', 'Sum','#,##0.00_ ;-#,##0.00'],
	['TOTAL','SUMA TOTAL', 'Sum', '#,##0.00_ ;-#,##0.00']],'Resumen_Ej','Total_por_Ejecutivo',2)



def pivot_table(wb: object, ws1: object, pt_ws: object, ws_name: str, pt_name: str, pt_rows: list, pt_filters: list, pt_fields: list, column: int):

    # pivot table location
	pt_loc = len(pt_filters) + 2
    
    # grab the pivot table source data
	pc = wb.PivotCaches().Create(SourceType=win32c.xlDatabase, SourceData=ws1.UsedRange)
    
    # create the pivot table object
	pc.CreatePivotTable(TableDestination=f'{ws_name}!R{pt_loc}C{column}', TableName=pt_name)

    # selecte the pivot table work sheet and location to create the pivot table
	pt_ws.Select()
	pt_ws.Cells(pt_loc, 1).Select()	

    # Sets the rows, columns and filters of the pivot table

	for field_list, field_r in ((pt_filters, win32c.xlPageField), 
                                (pt_rows, win32c.xlRowField)):
		for i, value in enumerate(field_list):
			pt_ws.PivotTables(pt_name).PivotFields(value).Orientation = field_r
			pt_ws.PivotTables(pt_name).PivotFields(value).Position = i + 1

    # Sets the Values of the pivot table
	for field in pt_fields:
		pt_ws.PivotTables(pt_name).AddDataField(pt_ws.PivotTables(pt_name).PivotFields(field[0]), field[1], field[2]).NumberFormat = field[3]

    # Visiblity True or Valse
	pt_ws.PivotTables(pt_name).ShowValuesRow = True
	pt_ws.PivotTables(pt_name).ColumnGrand = True



def run_excel(f_path : str, sheet_name: str, pt_filters: list,pt_rows : list, pt_fields : list, pt_name :str, sheet_pt : str, col: int = 1):
	filename = Path(f_path)

    # create excel object
	excel = win32.gencache.EnsureDispatch('Excel.Application')

    # excel can be visible or not
	excel.Visible = True  # False
    
    # try except for file / path
	try:
		wb = excel.Workbooks.Open(filename)
	except com_error as e:
		if e.excepinfo[5] == -2146827284:
			print(f'Failed to open spreadsheet.  Invalid filename or location: {filename}')
		else:
			raise e
		sys.exit(1)

    # set worksheet

	aggregations = {'Sum': win32c.xlSum, 'Avg': win32c.xlAverage, 'Cnt': win32c.xlCount}	
	pt_fields = [[ aggregations.get(element)if index == 2 else element for index,element in enumerate(list_) ] for list_ in pt_fields]

	ws1 = wb.Sheets(sheet_name)
    
    # Setup and call pivot_table
	wb.Sheets.Add().Name = sheet_pt
	ws2 = wb.Sheets(sheet_pt)
    
	pivot_table(wb, ws1, ws2, sheet_pt, pt_name, pt_rows, pt_filters, pt_fields,col)

	wb.Close(True)
	excel.Quit()


if __name__ == '__main__':
	main()

	


