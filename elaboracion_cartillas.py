import pandas as pd
from sqlalchemy import create_engine, text
import pygsheets as gc

def main():

    connection_str = 'mysql+pymysql://reports:cognos@192.168.1.238/mayaprin?charset=utf8'

    db_connection = create_engine(connection_str)

    query_str = """SELECT j_number AS OP,job200.j_customer AS Cliente, j_rep AS Vendedor, CONCAT(j_title1, IFNULL(j_title2,'')) AS Trabajo
    FROM job200
    WHERE (job200.j_tech_spec LIKE '%se agregan%pliegos%cartilla%' OR job200.j_tech_spec LIKE '%genera%cartilla%' OR job200.j_tech_spec LIKE '%Renovar cartilla%' OR job200.j_tech_spec LIKE '%REALIZAR CARTILLA%')
    AND job200.j_status NOT IN ('X')
    AND job200.j_booked_in >= '2021-11-01'"""

    df_cartillas = pd.read_sql_query(text(query_str), con=db_connection)

    json_auth_path = "C://Users//User//Documents//Analisis  Desarrollo Costos//Scripts//Python//secure_path//gcpApikey.json"

    gcon = gc.authorize(service_file = json_auth_path)

    wsheet = gcon.open_by_url('https://docs.google.com/spreadsheets/d/1Bv6ql9io2K_n_YW1D7rHWxpD3mz39Yuf5yIbdH2o8qU/edit#gid=1530866914')

    sheet = wsheet.worksheet_by_title('Control de entrega')

    sheet.clear(start='A2', end=f'D{sheet.rows}')

    sheet.set_dataframe(df_cartillas, (2,1), copy_head = False)

if __name__ == '__main__':
    main()
