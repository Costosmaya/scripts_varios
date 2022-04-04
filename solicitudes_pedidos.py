import pandas as pd
from sqlalchemy import create_engine, text
import pygsheets as gc

def main():

    connection_str = 'mysql+pymysql://reports:cognos@192.168.1.238/mayaprin?charset=utf8'
    json_auth_path = "C://Users//User//Documents//Analisis  Desarrollo Costos//Scripts//Python//secure_path//gcpApikey.json"
    gcon = gc.authorize(service_file = json_auth_path)

    wsheet = gcon.open_by_url('https://docs.google.com/spreadsheets/d/1YN6FiDWjmd1Rbsp3qzPkkCyX1br7DrsIRRpBOPejsaw/edit#gid=1370486329')

    sheet = wsheet.worksheet_by_title('Solicitud Pedidos')

    query_str = """SELECT job200.j_number AS Pedido, wo200.wo_number AS 'OT', req.item, itm.name AS Descripcion, SUM(req.qty_req) AS Cantidad, SUM(req.req_total_cost) AS Costo,
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
        WHEN itm.sl_analysis <> 'EMPAQUES' then 'revisado'
        ELSE
        'empaque'
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
        WHEN itm.sl_analysis <> 'EMPAQUES' then 6
        ELSE
        7
    END AS num_proceso
        FROM job200
        INNER JOIN wo200 ON job200.j_number = wo200.wo_job
        INNER JOIN wo_task200 ON wo200.wo_number = wo_task200.tk_wonum
        INNER JOIN req ON wo_task200.tk_id = req.req_task_id
        INNER JOIN itm ON req.item = itm.code
        WHERE job200.j_booked_in BETWEEN '2021-12-01' AND CURDATE()
        AND job200.j_status <> 'X'
        AND itm.code NOT LIKE 'K-%'
        AND itm.code NOT LIKE 'C-PRUEBA%'
        AND itm.code NOT IN ('VA-FABR-TROQUEL','X-FAB-CLICHE ESTAMPA')
        GROUP BY job200.j_number, wo200.wo_number, req.item, itm.name
    """
    db_connection = create_engine(connection_str)

    df_solicitudes = pd.read_sql_query(text(query_str),con=db_connection)

    sheet.clear(start='A2')
    sheet.set_dataframe(df_solicitudes, (2,1), copy_head=False)

if __name__ == '__main__':
    main()