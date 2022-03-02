import time
import win32com.client as win32
from numpy import int64
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
from pathlib import Path
from pywintypes import com_error
import sys
from sqlalchemy import create_engine, text
import pandas as pd
import pygsheets as gc
from datetime import datetime, date

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

	query_str2 = """SELECT j.j_rep AS VENDEDOR, SUM(ist.ist_goods) AS 'SIN IVA', SUM(ist.ist_goods) * 0.12 AS 'IVA D'
	FROM job200 j
	INNER JOIN ist ON j.j_number = ist.ist_job
	INNER JOIN inv ON ist.ist_inv_id = inv.inv_id
	INNER JOIN charge ON ist.ist_id = charge.cg_ist_id
	WHERE inv.inv_customer NOT IN ('EXPROREP','RECICLAJE', 'LA JOYA', 'INKTRA', 'ITERMARKE')
	AND inv.inv_prefix <>'FSFA'
	AND charge.cg_type IN ('GEN','PRODLITHO')
	AND j.j_type NOT IN ('ADICIONAL')
	AND j.j_status <> 'X'
	AND MONTH(inv.inv_date) = MONTH(CURDATE())
	AND YEAR(inv.inv_date) = YEAR(CURDATE())
	GROUP BY j.j_rep
	ORDER BY 3 DESC;
"""

	df_pendientes = pd.read_sql_query(text(query_str), con=db_connection)

	df_Facturado = pd.read_sql_query(text(query_str2),con=db_connection)

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

	current_time = datetime.now()

	df_pendientes['Días desde Ingreso'] = current_time - df_pendientes['FECHA CREACION']

	df_pendientes['Días desde Reporte'] = current_time - df_pendientes['FECHA REPORTE']

	df_pendientes['Días desde Ingreso'] = df_pendientes['Días desde Ingreso'].astype('timedelta64[D]').fillna('')

	df_pendientes['Días desde Reporte'] = df_pendientes['Días desde Reporte'].astype('timedelta64[D]').fillna('')
	
	df_pendientes['FECHA REPORTE'] = df_pendientes['FECHA REPORTE'].apply(lambda x: '' if x == 'NaT' else x)

	df_pendientes['FECHA REPORTE'] = df_pendientes['FECHA REPORTE'].dt.strftime('%d/%m/%y').fillna('')

	

	df_pendientes['FECHA CREACION'] = df_pendientes['FECHA CREACION'].dt.strftime('%d/%m/%y')

	df_Facturado['TOTAL'] = df_Facturado['SIN IVA'] + df_Facturado['IVA D']

	print(df_pendientes)
	

	path = "c://Users//User//Documents//Prueba Pendientes.xlsx"


	with pd.ExcelWriter(path, engine='xlsxwriter') as writer:
		workbook = writer.book
		worksheet =workbook.add_worksheet('Detalle')

		currency_format = workbook.add_format()
		currency_format.set_num_format('#,##0.00_ ;-#,##0.00')


		text_format = workbook.add_format()
		text_format.set_text_wrap()

		(max_row, max_col) = df_pendientes.shape

		

		format_settings = [{'header':column,'header_format':text_format, 'format':currency_format if column in ['SIN IVA','TOTAL','Días desde Ingreso','Días desde Reporte'] else  workbook.add_format()} for i,column in enumerate(df_pendientes.columns)]

		worksheet.add_table(1,1,max_row, max_col,{'data':df_pendientes.values.tolist(),'style':'Table Style Medium 2','columns':format_settings})

		worksheet.set_column(0, max_col, 12)

		worksheet = workbook.add_worksheet('Facturado')

		currency_format2 = workbook.add_format()

		currency_format2.set_num_format('_-Q* #,##0.00_-;-Q* #,##0.00_-;_-Q* "-"??_-;_-@_-')
		
		(max_row2, max_col2) = df_Facturado.shape

		format_settings = [{'header':column,'header_format':text_format, 'format': workbook.add_format() if column not in ['SIN IVA','IVA D','TOTAL'] else currency_format2,
		'total_string':'Total' if column == 'VENDEDOR' else None, 'total_function': 'sum' if column != 'VENDEDOR' else None} for i,column in enumerate(df_Facturado.columns)
		 ]

		worksheet.add_table(1,1,max_row2+2, max_col2,{'data':df_Facturado.values.tolist(),'total_row':True,'style':'Table Style Medium 2','columns':format_settings})

		worksheet.set_column(0, max_col2, 16)
	
	run_excel(path, 'Detalle',['VENDEDOR'],['Estado Pedido','CLIENTE','OP','Días desde Ingreso','Días desde Reporte'], [['SIN IVA','Total SIN IVA', 'Sum','#,##0.00_ ;-#,##0.00'],
	['TOTAL','SUMA TOTAL', 'Sum', '#,##0.00_ ;-#,##0.00']],'Resumen','Resumen',2)

	time.sleep(20)

	run_excel(path, 'Detalle',[],['VENDEDOR','Estado Pedido'], [['SIN IVA','Total SIN IVA', 'Sum','#,##0.00_ ;-#,##0.00'],
	['TOTAL','SUMA TOTAL', 'Sum', '#,##0.00_ ;-#,##0.00']],'Resumen_Ej','Total_por_Ejecutivo',2)


	msg_body = """Hola, \n Adjunto reporte de facturación correspondiente a la fecha, quedo atento a cualquier comentario o duda. Saludos 
	
	
	
	
	
	\n -Este mensaje se envía automáticamente-"""

	send_mail('costos@gmail.com','costos@mayaprin.com',f'Facturacion - {date.today().strftime("%d/%m/%Y")}',msg_body,path,'smtp.gmail.com',587,
	'costos@mayaprin.com','Mayaprin100%')



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

def send_mail(send_from,send_to,subject,text,files,server,port,username='',password='',isTls=True):
    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = send_to
    msg['Date'] = formatdate(localtime = True)
    msg['Subject'] = subject
    msg.attach(MIMEText(text))

    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(files, "rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename=Pendientes Facturacion - {date.today().strftime("%d/%m/%Y")}.xlsx')
    msg.attach(part)

    smtp = smtplib.SMTP("smtp.gmail.com", port)
    if isTls:
        smtp.starttls()
    smtp.login(username,password)
    smtp.sendmail(send_from, send_to, msg.as_string())
    smtp.quit()


if __name__ == '__main__':
	main()

	


