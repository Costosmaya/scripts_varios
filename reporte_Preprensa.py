from sqlalchemy import create_engine, text
import pandas as pd


def main():
    db_connection_str = 'mysql+pymysql://reports:cognos@192.168.1.238/mayaprin'

    file_path = 'C://Users//User//Documents//prueba_preprensa.xlsx'

    db_connection = create_engine(db_connection_str)

    query_str = """SELECT tm_start AS Fecha,tm_op_code AS Operario, TIMEDIFF(tm_end,tm_start) AS Duracion, tm_non_prod AS Actividad, IFNULL(tm_note,'') AS nota
    FROM tm
    WHERE tm.tm_non_prod LIKE 'ARR%'
    AND MONTH(tm_start) = 1
    AND YEAR(tm_start) = YEAR(CURDATE())"""

    df_marcajes = pd.read_sql_query(text(query_str), con= db_connection)

    df_marcajes['Mes'] = df_marcajes['Fecha'].dt.month

    df_marcajes['Semana'] = df_marcajes['Fecha'].dt.isocalendar().week

    df_marcajes['Fecha'] = df_marcajes['Fecha'].dt.strftime('%d/%m/%y')

    df_marcajes['Duracion (h)'] = df_marcajes['Duracion'] / pd.Timedelta(hours = 1)

    df_marcajes = df_marcajes.round({'Duracion (h)':2})

    df_marcajes = df_marcajes[['Fecha','Mes','Semana','Operario','Duracion (h)','nota']]

    df_marcajes_mes = df_marcajes.copy().drop(columns=['Semana'])

    df_marcajes_mes = df_marcajes_mes.groupby(['Mes']).sum().reset_index()

    df_marcajes_mes['Mes'] = df_marcajes_mes['Mes'].astype(str)

    df_marcajes_mes['Mes'].replace({'1':'Enero','2':'Febrero','3':'Marzo','4':'Abril','5':'Mayo','6':'Junio','7':'Julio','8':'Agosto','9':'Septiembre','10':'Octubre','11':'Noviembre','12':'Diciembre'}, inplace=True)

    df_marcajes_mes['Costo'] =  df_marcajes_mes['Duracion (h)'] * 51.56

    df_marcajes_mes = df_marcajes_mes.round({'Costo':2})

    df_marcajes_semana = df_marcajes.copy().drop(columns=['Mes'])

    df_marcajes_semana = df_marcajes_semana.groupby(['Semana']).sum().reset_index()

    df_marcajes_semana['Costo'] =  df_marcajes_semana['Duracion (h)'] * 51.56

    df_marcajes_semana = df_marcajes_semana.round({'Costo':2})

    df_marcajes_diario = df_marcajes.copy().drop(columns=['Mes','Semana'])

    df_marcajes_diario = df_marcajes_diario.groupby(['Fecha']).sum().reset_index()

    df_marcajes_diario['Costo'] =  df_marcajes_diario['Duracion (h)'] * 51.56

    df_marcajes_diario = df_marcajes_diario.round({'Costo':2})



    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:

        workbook = writer.book

        format_currency = workbook.add_format()
        format_currency.set_num_format('#,##0.00_ ;-#,##0.00')

        text_format = workbook.add_format()
        text_format.set_text_wrap()

        worksheet = workbook.add_worksheet('Datos')

        (max_row, max_col) = df_marcajes.shape

        format_settings = [{'header':column,'header_format':text_format, 'format':text_format} for i,column in enumerate(df_marcajes.columns)]

        worksheet.add_table(1,1,max_row+1, max_col,{'data':df_marcajes.values.tolist(),'style':'Table Style Medium 2','columns':format_settings})

        worksheet.set_column(0, max_col, 12)

        worksheet = workbook.add_worksheet('Diario')

        (max_row, max_col) = df_marcajes_diario.shape

        format_settings = [{'header':column,'header_format':text_format, 'format':format_currency if column in ['Costo', 'Duracion (h)'] else text_format} for i,column in enumerate(df_marcajes_diario.columns)]

        worksheet.add_table(1,1,max_row+1, max_col,{'data':df_marcajes_diario.values.tolist(),'style':'Table Style Medium 2','columns':format_settings})

        worksheet.set_column(0, max_col, 12)

        worksheet = workbook.add_worksheet('Semanal')

        (max_row, max_col) = df_marcajes_semana.shape

        format_settings = [{'header':column,'header_format':text_format, 'format':format_currency if column in ['Costo', 'Duracion (h)'] else text_format} for i,column in enumerate(df_marcajes_semana.columns)]

        worksheet.add_table(1,1,max_row+1, max_col,{'data':df_marcajes_semana.values.tolist(),'style':'Table Style Medium 2','columns':format_settings})

        worksheet.set_column(0, max_col, 12)

        worksheet = workbook.add_worksheet('Mensual')

        (max_row, max_col) = df_marcajes_mes.shape

        format_settings = [{'header':column,'header_format':text_format, 'format':format_currency if column in ['Costo', 'Duracion (h)'] else text_format} for i,column in enumerate(df_marcajes_mes.columns)]

        worksheet.add_table(1,1,max_row+1, max_col,{'data':df_marcajes_mes.values.tolist(),'style':'Table Style Medium 2','columns':format_settings})

        worksheet.set_column(0, max_col, 12)










if __name__ == '__main__':
    main()
    