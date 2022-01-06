from sqlalchemy import create_engine, text
import pandas as pd
import pygsheets as gc
import numpy as np
import re
import sys

db_connection_str = 'mysql+pymysql://reports:cognos@192.168.1.238/mayaprin?charset=utf8'

db_connection = create_engine(db_connection_str)

query_str = """SELECT j.j_number AS OP,j.j_customer AS \'Cliente\',j.j_ordnum AS \'OC\',SUM(j.j_quantity) AS \'Cantidad Pedido\'	,SUM(trs.wt_good_qty) AS \'Cantidad Reportada\', Expedicion.Entregado AS \'Cantidad Expediciones\'
FROM job200 j
INNER JOIN wo200  ON j.j_number = wo200.wo_job
INNER JOIN wo_task200 tsk ON wo200.wo_number = tsk.tk_wonum
INNER JOIN wo_trans200 trs ON tsk.tk_id = trs.wt_task_id
INNER	JOIN (SELECT job.j_number AS jobnum, sum(dr.dr_quantity) AS Entregado
FROM job200 job
INNER JOIN wo200 ON job.j_number = wo200.wo_job
INNER	JOIN delreq_task_view dr ON wo200.wo_number = dr.dr_wonum
INNER JOIN delnote dl ON dl.dn_number = dr.dr_dnnum
GROUP BY job.j_number) AS Expedicion ON j.j_number = Expedicion.jobnum 
WHERE tsk.tk_code = \'REPORTADO\'
GROUP BY j.j_number;"""



query = pd.read_sql_query(text(query_str), con = db_connection)


query['Cantidad en Inventario'] = (query['Cantidad Reportada'] - query['Cantidad Expediciones'])

def ajustarCantidad(cantidad):
    if cantidad >0 :
        return cantidad
    else: 
        return 0

query['Cantidad en Inventario'] = query['Cantidad en Inventario'].apply(lambda x: ajustarCantidad(x))
json_auth_path = "C://Users//User//Documents//Analisis  Desarrollo Costos//Scripts//Python//secure_path//gcpApikey.json"
gcon = gc.authorize(service_file = json_auth_path)

sheet = gcon.open_by_url('https://docs.google.com/spreadsheets/d/1YN6FiDWjmd1Rbsp3qzPkkCyX1br7DrsIRRpBOPejsaw')


wsheet = sheet[2]

wsheet.clear(start='A2', end='G')

wsheet.set_dataframe(query, (1,1), nan= '')