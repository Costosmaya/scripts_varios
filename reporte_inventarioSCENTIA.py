from sqlalchemy import create_engine, text
import pandas as pd
import pygsheets as gc
import numpy as np
import re
import sys


def dataExtraction():

    db_connection_str = 'mysql+pymysql://reports:cognos@192.168.1.238/mayaprin?charset=utf8'

    db_connection = create_engine(db_connection_str)

    query_str = """SELECT j.j_number,j.j_ordnum as OC, 
    CASE
    WHEN j.j_status = \'C\' THEN \'Terminado\'
    ELSE \'Abierto\' END AS Estado,CONCAT(j.j_title1,IFNULL(j.j_title2,'')) AS Titulo,j.j_quantity, j.j_estval, ROUND(ist.ist_goods,2) AS Facturado, ROUND(ist.ist_goods/PUnit.PU,0) AS 'Unidades Facturadas'
,CASE
WHEN ist.ist_quantity = 1 THEN
ROUND(ist.ist_goods/PUnit.PU,0)
ELSE ist.ist_quantity END AS 'Unidades Facturadas Real', PUnit.PU
FROM job200 j
INNER JOIN ist ON j.j_number = ist.ist_job
INNER JOIN inv ON ist.ist_inv_id = inv.inv_id
INNER JOIN (SELECT job200.j_number AS op, (job200.j_estval/job200.j_quantity) AS PU
FROM job200) AS PUnit ON j.j_number = PUnit.op
WHERE j.j_customer = 'SCENTIA'
AND ist.ist_text NOT LIKE 'GASTOS ADMINISTRATIVOS'
AND YEAR(j.j_booked_in) >= 2021
AND j.j_type = 'CAJA'
AND MONTH(j.j_booked_in) >= 6;
"""

    query = pd.read_sql_query(text(query_str), con = db_connection)

    dfData = query[["j_number",'OC',"Estado","Titulo","j_quantity", "j_estval",'PU']]
    dfData.drop_duplicates(subset=["j_number"], inplace=True)

    dfSummarize = query[["j_number","Facturado","Unidades Facturadas","Unidades Facturadas Real"]].groupby(by=["j_number"]).sum()

    dfData = pd.merge(dfData, right=dfSummarize, how="inner", on="j_number")

    dfData['Unidades Pendientes'] = dfData['j_quantity'] - dfData['Unidades Facturadas Real']

    dfData = dfData[dfData['Unidades Pendientes'] != 0 ]

    json_auth_path = "C://Users//User//Documents//Analisis  Desarrollo Costos//Scripts//Python//secure_path//gcpApikey.json"
    
    gcon = gc.authorize(service_file = json_auth_path)

    sheet = gcon.open_by_url('https://docs.google.com/spreadsheets/d/1iTD921o_UnsTXjFN0PBkieMDSYtPmAVV1rIo9fWugyo')
    wsheet = sheet[0]

    wsheet.clear(start='A2', end='K1000')

    wsheet.set_dataframe(dfData, (2,1), copy_head = False, nan= '')

dataExtraction()