from sqlalchemy import create_engine, text
import pandas as pd
import pygsheets as gc

def main():
    db_connection_str = 'mysql+pymysql://reports:cognos@192.168.1.238/mayaprin?charset=utf8'

    query_str = """SELECT wo200.wo_job AS Pedido,wo_task200.tk_code AS Actividad, wo_trans200.wt_source_code AS 'Código Origen', IFNULL(staf.staf_name,itm.name) AS nombre, 
        wo_trans200.wt_started AS 'Fecha Marcaje',
        CASE
        WHEN ROUND(AVG(
        CASE
        WHEN charge.cg_status = 0 then 0
        ELSE 1
        END ),0) = 1 then 'Cerrado' ELSE 'Abierto' END AS Concepto
        FROM wo200
        INNER JOIN wo_task200 ON wo200.wo_number = wo_task200.tk_wonum
        INNER JOIN wo_trans200 ON wo_task200.tk_id = wo_trans200.wt_task_id
        LEFT JOIN staf ON wo_trans200.wt_source_code = staf.staf_code
        LEFT JOIN itm ON wo_trans200.wt_source_code = itm.code
        INNER JOIN charge ON wo200.wo_job = charge.cg_job
        WHERE wo_task200.tk_status = 'N'
        AND YEAR(wo_trans200.wt_started) = YEAR(CURDATE())
        GROUP BY wo200.wo_job, wo200.wo_number, wo_task200.tk_id
        ORDER BY wo200.wo_job ASC, wo_trans200.wt_started DESC;"""

    db_connection = create_engine(db_connection_str)

    df_marcajes = pd.read_sql_query(text(query_str), con=db_connection)
    
    json_auth_path = "C://Users//User//Documents//Analisis  Desarrollo Costos//Scripts//Python//secure_path//gcpApikey.json"
    gcon = gc.authorize(service_file = json_auth_path)
    
    wsheet = gcon.open_by_url('https://docs.google.com/spreadsheets/d/1I2ISrUS9UDgtKo4_j957dDbLp6yAoirXoMAVUDmP2iI/edit#gid=503348355')
    
    sheet = wsheet.worksheet_by_title('Marcajes Abiertos')

    sheet.clear(start='A2', end=f'D{sheet.rows}')

    sheet.set_dataframe(df_marcajes, (1,1))
    
if __name__ == '__main__':
    main()