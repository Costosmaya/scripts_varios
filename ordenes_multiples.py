import pandas as pd
from sqlalchemy import create_engine, text
import pygsheets as gc
import re
import os

from uritemplate.api import expand

def main():
    connection_str = 'mysql+pymysql://reports:cognos@192.168.1.238/mayaprin?charset=utf8'
    json_auth_path = "C://Users//User//Documents//Analisis  Desarrollo Costos//Scripts//Python//secure_path//gcpApikey.json"
    gcon = gc.authorize(service_file = json_auth_path)

    sql_str = """SELECT job200.j_number,CONCAT(job200.j_title1, CONCAT('',IFNULL(job200.j_title2,''))) AS Titulo, job200.j_ucode3 AS 'No. Motivos', job200.j_tech_spec
        FROM job200
        WHERE job200.j_ucode3 > 1
        AND DATE(job200.j_booked_in) >= '2022-02-23'"""

    db_connection = create_engine(connection_str)

    query = pd.read_sql_query(text(sql_str),con=db_connection)

    def find_values(value):
        motivos = int(value['No. Motivos'])
        matches = value['j_tech_spec'].strip().split('\n')[-motivos:]
        return ';'.join([string.strip() for  string in matches])
    
    query['matches'] = query.apply(find_values, axis=1)

    print(query)

    query= query[['j_number','Titulo','No. Motivos','matches']]

    max_motivos = int(query['No. Motivos'].max())+1

    columns = [str(num) for num in range(1,max_motivos)]

    df_motivos = query.copy()[['j_number','matches']]

    df_motivos[columns] = df_motivos['matches'].str.split(';', expand=True)

    df_motivos = df_motivos.melt(id_vars='j_number', value_vars=columns, var_name='No. Motivo',value_name='Motivo')

    df_desglose = pd.merge(query[['j_number','Titulo']],right=df_motivos[['j_number','Motivo']],how='left', on='j_number')

    df_desglose.dropna(inplace=True)

    df_desglose['Titulo Motivo'] = df_desglose['Titulo'] + ' ' + df_desglose['Motivo']

    df_desglose.drop(columns=['Titulo','Motivo'], inplace=True)

    df_desglose.to_csv(os.getcwd() +'/orders.csv')
    
    


    
if __name__ == '__main__':
    main()