import time
from venv import create
import win32com.client as win32
from numpy import int64
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
from pathlib import Path
from pywintypes import com_error
import sys
from sqlalchemy import create_engine, text
import pandas as pd
import pygsheets as gc
from datetime import datetime, date
from pendientes_facturacion import send_mail

def main():


    db_connection_str = 'mysql+pymysql://reports:cognos@192.168.1.238/mayaprin?charset=utf8'

    query_str = """SELECT j.j_number AS OP, j.j_customer AS Cliente, itm.name AS Descripcion,itm.fvals_2 AS Gramos, req.qty_req AS Cantidad
    FROM job200 j
    INNER JOIN wo200 ON j.j_number = wo200.wo_job
    INNER JOIN wo_task200 tk ON wo200.wo_number = tk.tk_wonum
    INNER JOIN req ON tk.tk_id = req.req_task_id
    INNER JOIN itm ON req.item = itm.code
    WHERE itm.class_code = 'PAPELHOJA'
    AND DATE(j.j_booked_in) = CURDATE();"""

    db_connection = create_engine(db_connection_str)

    df_ingresos = pd.read_sql_query(text(query_str), con=db_connection)

    path = 'c://Users//User//Documents//Reporte Ingresos - Papel.xlsx'

    with pd.ExcelWriter(path, engine='xlsxwriter') as writer:
        worbook = writer.book

        worksheet = worbook.add_worksheet('Datos')

        (max_row, max_col) = df_ingresos.shape

        column_settings = [{'header': column, 'format': worbook.add_format().set_text_wrap()} for column in df_ingresos.columns]

        worksheet.add_table(1,1,max_row, max_col,{'data':df_ingresos.values.tolist(),'style':'Table Style Medium 2','columns':column_settings})
        worksheet.set_column(0, max_col, 12)

        msg_body = """Hola Eugenia, \n Adjunto reporte de Ingresos con detalle de papel correspondiente a la fecha, quedo atento a cualquier comentario o duda. Saludos  \n -Este mensaje se envía automáticamente-"""

    send_mail('costos@mayaprin.com','costos@mayaprin.com',f'Ingresos - Papel {date.today().strftime("%d/%m/%Y")}',msg_body,path,'smtp.gmail.com',587,
    'costos@mayaprin.com','Mayaprin100%')



if __name__ == '__main__':
    main()
