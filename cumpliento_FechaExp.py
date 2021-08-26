from sqlalchemy import create_engine, text
import pandas as pd
import pygsheets as gc

db_connection_str = 'mysql+pymysql://reports:cognos@192.168.1.238/mayaprin'

db_connection = create_engine(db_connection_str)

query_str = """SELECT j.j_number, dr.dr_dnnum, dl.dn_despatched
FROM job200 j
INNER JOIN wo200 ON j.j_number = wo200.wo_job
INNER	JOIN delreq_task_view dr ON wo200.wo_number = dr.dr_wonum
INNER JOIN delnote dl ON dl.dn_number = dr.dr_dnnum
ORDER BY dl.dn_despatched;"""

query = pd.read_sql_query(text(query_str), con=db_connection)   

json_auth_path = "C://Users//User//Documents//Analisis  Desarrollo Costos//Scripts//Python//secure_path//gcpApikey.json"
gcon = gc.authorize(service_file = json_auth_path)

sheet = gcon.open_by_url('https://docs.google.com/spreadsheets/d/1rdzDkynYglw1fpPAzGJM3-ac1XRTYxZZAwOjjAQpUP0')

wsheet = sheet[8]

wsheet.clear(start='A2')

wsheet.set_dataframe(query, (2,1), copy_head = False, nan= '')

print('done')
