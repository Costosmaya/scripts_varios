import pandas as pd
from sqlalchemy import create_engine, text
import pygsheets as gc


def main():
    connection_str = 'mysql+pymysql://reports:cognos@192.168.1.238/mayaprin?charset=utf8'

    json_auth_path = "C://Users//User//Documents//Analisis  Desarrollo Costos//Scripts//Python//secure_path//gcpApikey.json"
    
    gcon = gc.authorize(service_file = json_auth_path)

    wsheet = gcon.open_by_url('https://docs.google.com/spreadsheets/d/1YN6FiDWjmd1Rbsp3qzPkkCyX1br7DrsIRRpBOPejsaw/edit#gid=1139174598')

    sheet = wsheet.worksheet_by_title('NCI-Consumos')

    query_str = """SELECT j.j_number AS OP, wo200.wo_number as OT, itm.code AS Codigo,CONCAT(itm.name,CONCAT(IFNULL(itm.itm_name_2,''),IFNULL( itm.itm_name_3,''))) AS Descripcion,
     SUM(iss.iss_quantity) AS Cantidad,SUM(iss.iss_value) AS Valor,
    CASE
        WHEN itm.class_code = 'PAPELHOJA' then 'papel'
        WHEN itm.class_code = 'PLACAS'
        OR (itm.class_code = 'TINTA' AND itm.sl_analysis <> 'BARNICES')
        OR itm.class_code = 'CLICK'
        OR itm.class_code = 'TINTA_ESP' then 'Impreso'
        WHEN itm.sl_analysis = 'BARNICES' then 'barnizado'
        WHEN itm.class_code = 'PLASTICOS' then 'laminado'
        WHEN itm.class_code = 'FOIL' then 'estampado'
        WHEN itm.class_code = 'TROQUEL' then 'troquelado'
        WHEN itm.name LIKE 'PEGAMENTO%' then 'pegado'
        WHEN itm.sl_analysis <> 'EMPAQUE' then 'armado'
        ELSE
        'fin'
    END AS Proceso,
    CASE
        WHEN itm.class_code = 'PAPELHOJA' then 0
        WHEN itm.class_code = 'PLACAS'
        OR (itm.class_code = 'TINTA' AND itm.sl_analysis <> 'BARNICES')
        OR itm.class_code = 'CLICK'
        OR itm.class_code = 'TINTA_ESP' then 1
        WHEN itm.sl_analysis = 'BARNICES' then 2
        WHEN itm.class_code = 'PLASTICOS' then 2
        WHEN itm.class_code = 'FOIL' then 3
        WHEN itm.class_code = 'TROQUEL' then 4
        WHEN itm.name LIKE 'PEGAMENTO%' then 5
        WHEN itm.sl_analysis <> 'EMPAQUE' then 5
        ELSE
        6
    END AS num_proceso
    FROM job200 j
    INNER JOIN wo200 ON j.j_number = wo200.wo_job
    INNER JOIN wo_task200 tsk ON wo200.wo_number = tsk.tk_wonum
    INNER JOIN iss ON tsk.tk_id = iss.iss_task_id
    INNER JOIN itm ON iss.item = itm.code
    WHERE iss.id NOT IN  (SELECT iss.id FROM iss INNER JOIN itm ON iss.item = itm.code WHERE itm.class_code <> 'PAPELHOJA'
	 AND (iss.note LIKE '%Reimp%' OR iss.note LIKE '%REIMP%' OR iss.note LIKE '%reimp%') AND iss.when_issued BETWEEN '2021-12-01' AND CURDATE())
    AND iss.when_issued BETWEEN '2021-12-01' AND CURDATE()
    GROUP BY j.j_number,wo200.wo_number, itm.code, iss.id"""

    db_connection = create_engine(connection_str)

    df_reprocesos = pd.read_sql_query(text(query_str),con=db_connection)

    sheet.clear(start='A2', end=f'G{sheet.rows}')

    sheet.set_dataframe(df_reprocesos, (2,1), copy_head = False)

if __name__ == '__main__':
    main()