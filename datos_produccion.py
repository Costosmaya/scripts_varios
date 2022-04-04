from sqlalchemy import create_engine, text
import pandas as pd
import pygsheets as gc
import numpy as np
import re
import sys


def main():

    db_connection_str = 'mysql+pymysql://reports:cognos@192.168.1.238/mayaprin?charset=utf8'

    db_connection = create_engine(db_connection_str)

    db_connection.execute(text('''SET lc_time_names = 'es_MX';'''));

    query_str = '''SELECT MONTHNAME(wo_trans200.wt_started) as Mes, cc.cc_name AS 'Centro Coste',act.act_code AS Actividad,act.act_name,
        CASE
        WHEN act.act_name LIKE '%fondo%grande%' THEN 'Fondo Grande'
        WHEN act.act_name LIKE '%fondo%mediano%' THEN 'Fondo Mediano'
        WHEN act.act_name LIKE '%fondo%pequeño' THEN 'Fondo Pequeño'
        WHEN act.act_name LIKE '%lateral%grande%' THEN 'Lateral Grande'
        WHEN act.act_name LIKE '%lateral%pequeño%' THEN 'Lateral Pequeño'
        WHEN act.act_name LIKE '%lateral%mini%' THEN 'Lateral Mini'
        ELSE act.act_name END AS Tipo, SUM(wo_trans200.wt_good_qty + wo_trans200.wt_bad_qty) AS Unidades
        FROM job200
        INNER JOIN wo200 ON job200.j_number = wo200.wo_job
        INNER JOIN wo_task200 ON wo200.wo_number = wo_task200.tk_wonum
        INNER JOIN wo_trans200 ON wo_task200.tk_id = wo_trans200.wt_task_id
        INNER JOIN act ON wo_task200.tk_code = act.act_code
        INNER JOIN res ON wo_trans200.wt_resource = res.res_code
        INNER JOIN cc ON res.res_cc = cc.cc_code
        WHERE job200.j_type = 'CAJA'
        AND YEAR(wo_trans200.wt_started) >= YEAR(CURDATE())
        AND wo_trans200.wt_source = 'TS'
        AND wo_trans200.wt_resource LIKE 'PEG CAJ%'
        AND act.act_analysis = 'TIR'
        AND (act.act_code NOT LIKE '%AUX%' AND act.act_code NOT LIKE '%REV%')
        GROUP BY MONTH(wo_trans200.wt_started),cc.cc_name,act.act_code;'''

    query = pd.read_sql_query(text(query_str), con = db_connection)

    df_agrupado = query.groupby(['Mes','Tipo']).sum().reset_index()

    json_auth_path = "C://Users//User//Documents//Analisis  Desarrollo Costos//Scripts//Python//secure_path//gcpApikey.json"
    
    gcon = gc.authorize(service_file = json_auth_path)

    sheet = gcon.open_by_url('https://docs.google.com/spreadsheets/d/1I2ISrUS9UDgtKo4_j957dDbLp6yAoirXoMAVUDmP2iI/edit#gid=0')
    wsheet = sheet.worksheet_by_title('Pegue cajas')

    wsheet.set_dataframe(df_agrupado, (2,1), copy_head = False, nan= '')

if __name__ == '__main__':
    main()