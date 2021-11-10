from sqlalchemy import create_engine, text
import pandas as pd
import pygsheets as gc

db_connection_str = 'mysql+pymysql://reports:cognos@192.168.1.238/mayaprin'

db_connection = create_engine(db_connection_str)

query_str = """SELECT job200.j_number, job200.j_rep AS Vendedor,SUM(charge.cg_value) AS 'Concepto Facturable', IFNULL(SUM(ist.ist_goods),-9999.99) AS 'Total Facturado' FROM 
job200
INNER JOIN charge ON charge.cg_job = job200.j_number
LEFT JOIN ist ON job200.j_number = ist.ist_job
LEFT JOIN inv ON ist.ist_inv_id = inv.inv_id

WHERE YEAR(charge.cg_date_created) = 2021
GROUP BY job200.j_number;"""

query = pd.read_sql_query(text(query_str), con=db_connection)   

json_auth_path = "C://Users//User//Documents//Analisis  Desarrollo Costos//Scripts//Python//secure_path//gcpApikey.json"
gcon = gc.authorize(service_file = json_auth_path)

sheet = gcon.open_by_url('https://docs.google.com/spreadsheets/d/1OanYoy-Rabi2QAlcm_gd0VkrJ86bkkDlxCQy244BR40')

wsheet = sheet[0]

wsheet.clear(start='A2')

wsheet.set_dataframe(query, (2,1), copy_head = False, nan='')

print('done')
