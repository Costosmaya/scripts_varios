from sqlalchemy import create_engine, text
import pandas as pd
import pygsheets as gc

db_connection_str = 'mysql+pymysql://reports:cognos@192.168.1.238/mayaprin'

db_connection = create_engine(db_connection_str)

query_str = """SELECT job200.j_number, job200.j_ucode6
FROM job200
WHERE YEAR(job200.j_booked_in) = YEAR(CURDATE());"""

query = pd.read_sql_query(text(query_str), con=db_connection)   

json_auth_path = "C://Users//User//Documents//Analisis  Desarrollo Costos//Scripts//Python//secure_path//gcpApikey.json"
gcon = gc.authorize(service_file = json_auth_path)

sheet = gcon.open_by_url('https://docs.google.com/spreadsheets/d/1YN6FiDWjmd1Rbsp3qzPkkCyX1br7DrsIRRpBOPejsaw')

wsheet = sheet[0]

wsheet.clear(start='A2')

wsheet.set_dataframe(query, (2,1), copy_head = False, nan='')

print('done')
