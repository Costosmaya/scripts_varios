# To add a new cell, type '# %%'
# To add a new markdown cell, type '# %% [markdown]'
# %%
from sqlalchemy import create_engine, text
from openpyxl import load_workbook
import pandas as pd
import numpy as np
import re
import sys

db_connection_str = 'mysql+pymysql://reports:cognos@192.168.1.238/mayaprin?charset=utf8'

db_connection = create_engine(db_connection_str)

#query mensual
query_str = """SELECT Distinct job200.j_number, job200.j_orig,e4e.ee_hdrnum, e4e.ee_estnum, sect_text, item.itm_code,
stkitm.itm_fvals_0 AS Ancho, stkitm.itm_fvals_1 AS Alto, sum(iss.iss_value) AS iss_value, sum(iss.iss_quantity) AS iss_quantity
FROM  e4e 
INNER JOIN est2wo  ON (e4e.ee_hdrnum = est2wo.ew_hdrnum AND e4e.ee_estnum = est2wo.ew_estnum)
INNER	JOIN wo200 ON est2wo.ew_wonum = wo200.wo_number
INNER JOIN job200 ON wo200.wo_job = job200.j_number
INNER JOIN esttext ON(e4e.ee_persistid = esttext.id)
INNER JOIN wo_task200 tsk ON wo200.wo_number = tsk.tk_wonum
INNER JOIN iss ON iss.iss_task_id = tsk.tk_id
inner join itm_cls_view item ON iss.item = item.itm_code
INNER JOIN stkitm ON item.itm_code = stkitm.itm_code
WHERE esttext.sect_num = 0
AND item.itm_is_paper = 1
AND YEAR(iss.when_issued) = YEAR(CURDATE())
AND MONTH(iss.when_issued) = MONTH(CURDATE())
GROUP BY  job200.j_number, job200.j_orig, e4e.ee_hdrnum, e4e.ee_estnum, item.itm_code;"""



query = pd.read_sql_query(text(query_str), con = db_connection)
query


# %%



columns_finishedDf = ["j_number","j_orig","ee_hdrnum","ee_estnum","itm_code","Ancho","Alto","iss_value","iss_quantity","workingWidth", "workingDepth","numberOutOfSheet"
,"imposedWidth","imposedDepth",	"Area Almacen","costo Unitario","Costoxinch2", "Area Prensa", "Area Imposicion", "Area desperdiciada", "% Area desperdiciada", "Costo Desperdicio"]

dfCostos = pd.DataFrame(columns=columns_finishedDf)
orders = list(dict.fromkeys(query["j_number"].tolist()))
for order in orders:  
    subdf1 = query.loc[query.j_number == order]
    headers = list(dict.fromkeys(subdf1["ee_hdrnum"].tolist()))
    for header in headers:        
        subdf2 = subdf1.loc[subdf1.ee_hdrnum == header]
        estimates = list(dict.fromkeys(subdf2["ee_estnum"].tolist()))
        for estimate in estimates:            
            subdf = subdf2.loc[subdf2.ee_estnum == estimate]
            dicto = subdf["sect_text"].tolist()
            matches1 = re.finditer("\'matCode\'", dicto[0])
            matches1_start = [match.start() for match in matches1]

            attributes = ("workingWidth", 'workingDepth', 'numberOutOfSheet', 'imposedWidth', 'imposedDepth')

            materials = []
            for index in matches1_start:
                materials.append(dicto[0][index:].split(":")[1].split(",")[0].replace("\'","").strip())
            subdf = subdf.loc[subdf.itm_code.isin(materials)]
            materials = subdf["itm_code"]
            total = len(subdf.index) 
            for attribute in attributes:               
                values = []                
                matches = re.finditer(attribute, dicto[0])
                matches_start = [match.start() for match in matches]            
                for x in range(len(matches_start)):                    
                    index = matches_start[x]
                    str_ = dicto[0][index:].split(":")[1].split(",")[0].replace("\'","").strip()
                    value = int(re.search(r'\d+', str_).group())
                    if(value == 0 or (value< 25 and attribute != "numberOutOfSheet")):
                        continue
                    else:
                        if(len(values) < total): 
                            values.append(value)

                subdf[attribute] = pd.to_numeric(values)


            subdf.drop(columns=["sect_text"], inplace=True)
        
            def convert_toInch(data):
                return round(data/25.4,2)

            subdf[["workingWidth", "workingDepth","imposedWidth", "imposedDepth"]] = subdf[["workingWidth", "workingDepth","imposedWidth", "imposedDepth"]].apply(convert_toInch)

            subdf[["Ancho", "Alto"]] = subdf[["Ancho", "Alto"]].apply(lambda x: round(x/25.4))

            subdf["Area Almacen"] = subdf["Ancho"] * subdf["Alto"]

            subdf["costo Unitario"] = subdf["iss_value"]/subdf["iss_quantity"]

            subdf["Costoxinch2"] = subdf["costo Unitario"]/subdf["Area Almacen"]

            subdf["Area Prensa"] = subdf["workingWidth"] * subdf["workingDepth"]

            subdf["Area Imposicion"] = subdf["imposedWidth"] * subdf["imposedDepth"]

            subdf["Area desperdiciada"] = (subdf["Area Almacen"] - (subdf["Area Prensa"] * subdf["numberOutOfSheet"])) + (subdf["Area Prensa"] - subdf["Area Imposicion"])
            
            subdf["% Area desperdiciada"] = subdf["Area desperdiciada"]/subdf["Area Almacen"]

            subdf["Costo Desperdicio"] = subdf["Area desperdiciada"] * subdf["Costoxinch2"] * subdf["numberOutOfSheet"]*subdf["iss_quantity"]



            dfCostos = dfCostos.append(subdf, ignore_index = True)


# %%
dfCostoExtra = dfCostos.copy()

dfCostoExtra = dfCostoExtra[dfCostoExtra["% Area desperdiciada"]>0.08]

dfCostoExtra["% Area extra desperdiciada"] = dfCostoExtra["% Area desperdiciada"] - 0.08

dfCostoExtra["Area extra desperdiciada"] = dfCostoExtra["Area Almacen"] * dfCostoExtra["% Area extra desperdiciada"]

dfCostoExtra["Costo desperdicio extra"] = dfCostoExtra["Area extra desperdiciada"] * dfCostoExtra["Costoxinch2"] * dfCostoExtra["iss_quantity"]



# %%
dfResumen = dfCostoExtra.copy()

dfResumen = dfResumen.groupby("j_orig")[["Costo desperdicio extra","iss_value"]].sum()

dfResumen["% Costo desperdicio vs Costo Consumo"] = dfResumen["Costo desperdicio extra"]/dfResumen["iss_value"]

dfresumen = dfResumen.rename(columns={'iss_value':'Costo Consumos'})


# %%

path = "c://Users//User//Documents//Costo Desperdicios_datos.xlsx"
with pd.ExcelWriter(path, engine='xlsxwriter') as writer:

    dfResumen.to_excel(writer, sheet_name="Resumen");
    
    dfCostos.to_excel(writer, sheet_name="Datos", index = False)

    dfCostoExtra.to_excel(writer, sheet_name="Costo Extra", index= False)



    workbook = writer.book
    
    worksheet = writer.sheets['Costo Extra']

    (max_row, max_col) = dfCostoExtra.shape

    wrap_format     = workbook.add_format({'text_wrap': 1})

    column_settings = [{'header': column, 'header_format':wrap_format} for column in dfCostoExtra.columns]


    worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings})

    # Make the columns wider for clarity.
    worksheet.set_column(0, max_col - 1, 12)