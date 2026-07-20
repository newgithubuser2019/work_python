# IMPORTS
import datetime
# import decimal
# import json
# import math
import os
# import pprint
# import re
# import shutil
import sys
from functools import reduce

import connectorx as cx
# import df2tables
import mssql_python
import numpy as np
import openpyxl
import pandas as pd
import plotly.express as px
import polars as pl
from dateutil.relativedelta import relativedelta
# from itables.streamlit import interactive_table
# from pandas.tseries.offsets import DateOffset
# from sqlalchemy import create_engine

import _my_functions

"""
pd.set_option("display.max_rows", 1500)
pd.set_option("display.max_columns", 100)
pd.set_option("max_colwidth", 30)
pd.set_option("expand_frame_repr", False)
"""
# ---------------------------------------------------------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------ global variables
USERPROFILE = os.environ["USERPROFILE"]

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------ prompts for user input
prompt1 = "\nгод отчетного периода: "
prompt2 = "\nмесяц отчетного периода: "
prompt3 = "\nДля продолжения необходимо загрузить договоры с затертыми цепочками и новые договоры. Выполнена ли загузка? "

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------ user inputs
inp1 = input(prompt1)
inp2 = input(prompt2)

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------ file paths
# path_1 = USERPROFILE + "\\Documents\\Работа\\отчетность\\ежедневно\\накопительный отчет\\"
# listoffiles_1 = os.listdir(path_1)

# ------------------------------------------------------------------ file names
if len(str(inp2)) == 2:
        month = str(inp2)
else:
    month = "0" + str(inp2)
filename1 = "P:\\Documents\\ДБ\\СРП\\Компенсации и льготы\\Потапов Д\\отчеты\\разное\\договоры для коэфф расторжений\\" + str(inp1) + "\\" + month + "\\договоры.xlsx"
filename2 = "P:\\Documents\\ДБ\\СРП\\Компенсации и льготы\\Потапов Д\\отчеты\\контакт\\126 отчет\\" + str(inp1) + "\\" + month + "\\126.xlsx"
filename3 = USERPROFILE + "\\Documents\\output.xlsx"
filename4 = "P:\\Documents\\ДБ\\СРП\\Компенсации и льготы\\Потапов Д\\отчеты\\ЗУП\\история изменений ФИО\\" + str(inp1) + "\\" + month + "\\история изм ФИО.xlsx"

# ------------------------------------------------------------------ database URIs
# sql_username = getpass.getuser()
# sql_password = getpass.getpass(prompt="SQL password: ", stream=None)
# sql_password = getpass.getpass(prompt="SQL password: ", stream=None, echo_char="*") # echo_char parameter added in python 3.14
# uri1 = "mssql://" + sql_username + ":" + sql_password + "@vls-sql-zup-dev:1433/HR_CAB"
# print(uri1)
polars_uri_1 = "mssql://vls-sql-zup-dev:1433/HR_CAB?trusted_connection=true"
mssql_python_uri_1 = "Server=vls-sql-zup-dev,1433;Database=HR_CAB;Trusted_connection=yes;TrustServerCertificate=yes"

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------ даты и периоды
bonus_period = datetime.date(int(inp1), int(inp2), 1)
# print(bonus_period.month)
# print(bonus_period)
list_of_12_periods = []
for i in range (1,13):
    past_period = bonus_period - relativedelta(months=i)
    # print(past_period)
    # print(calendar.month_name[past_period.month])
    list_of_12_periods.append("\'" + str(past_period) + "\'")
# print(list_of_12_periods)
previous_12_periods = ", ".join(list_of_12_periods)
# print(previous_12_periods)
cutoff_termination_date = datetime.date(int(inp1), int(inp2) + 1, 5)
# print(cutoff_termination_date)

# ------------------------------------------------------------------ проверка базы договоров на наличие новых договоров
query_p1 = "SELECT * FROM dbo.BONUSDogBase_python WHERE dogovor_period = "
past_period = bonus_period - relativedelta(months=1)
query_p2 = "\'" + str(past_period) + "\'"
query = query_p1 + query_p2
# print(query)
df_from_sql = pl.read_database_uri(
    uri=polars_uri_1,
    query=query,
    engine="connectorx",
    # schema_overrides={
        #"Год": pl.Int64,
        #},
    )
# print(df_from_sql.shape[0])

# ------------------------------------------------------------------ выборка расторгнутых договоров
if df_from_sql.shape[0] == 0:
    # print("\nНе загружены новые договоры / договоры с затертыми цепочками")

    df_from_excel = pl.read_excel(
        # engine="openpyxl",
        schema_overrides={
            "Дата удержания ": pl.Date,
            "LoadDt": pl.Date,
            "dogovor_period": pl.Date,
            },
        source=filename1,
        sheet_name="Лист1",
        has_header=True
        # columns="A:B"
    )
    # print(df_from_excel.head())
    df_rastorgnut = df_from_excel.filter((pl.col("Комментарий") == "расторжение"))
    # df_new = df_from_excel.filter((pl.col("Комментарий") != "расторжение") | pl.col("Комментарий").is_null())
    # print(df_rastorgnut)

    # ------------------------------------------------------------------ удаление расторгнутых договоров из базы
    if df_rastorgnut.shape[0] > 0:
        rastorgn_dogovory = df_rastorgnut["Номер договора"].to_list()
        # print(rastorgn_dogovory)
        rastorgn_dogovory_query = []
        for i in rastorgn_dogovory:
            rastorgn_dogovory_query.append("N" + "\'" + i + "\'")
        # print(rastorgn_dogovory_query)

        query_p1 = "DELETE FROM dbo.BONUSDogBase_python WHERE [Номер договора] in ("
        query_p2 = ", ".join(rastorgn_dogovory_query)
        query_p3 = ")"
        query = query_p1 + query_p2 + query_p3
        # print(query)
        # sys.exit()

        # apache arrow adbc
        """
        from adbc_driver_manager import dbapi
        with dbapi.connect(polars_uri_1) as conn:
            with conn.cursor() as cur:
                cur.execute("query")
                conn.commit()
                print("Rows deleted:", cur.rowcount)
        """

        # mssql-python
        # conn_str = "Server=vls-sql-zup-dev,1433;Database=HR_CAB;Trusted_connection=yes;TrustServerCertificate=yes"
        with mssql_python.connect(mssql_python_uri_1) as conn:
            with conn.cursor() as cur:
                cur.execute(query)
                conn.commit()
                print("\nRows deleted:", cur.rowcount)
        print("Удалены договоры, по которым затерты цепочки")

    # ------------------------------------------------------------------ загрузка новых договоров в базу
    """
    inp3 = input(prompt3)
    if inp3.lower() == "да":
        pass
    else:
        print("проверка сработала")
        sys.exit()
    # print(df_from_excel.shape[0])
    # print(df_from_excel)
    """

    # apache arrow adbc
    """
    df_from_excel.write_database(
                    table_name="BONUSDogBase_python",
                    connection=polars_uri_1,
                    engine="adbc",
                    if_table_exists="append"
                )
    """

    # sqlalchemy
    """
    db_url = ("mssql+pyodbc://vls-sql-zup-dev:1433/HR_CAB?trusted_connection=true&driver=SQL+Server&Encrypt=no")
    engine = create_engine(
        db_url,
        # fast_executemany=True, # uses a lot of ram
        use_setinputsizes=False,
        isolation_level="AUTOCOMMIT"
        )
    with engine.connect() as connection:
        with connection.begin():
            df_from_excel.write_database(
                table_name="BONUSDogBase_python",
                connection=connection,
                # engine="adbc",
                if_table_exists="append"
            )
            # connection.commit()
    print("Загружены договоры с затертыми цепочками и новые договоры")
    """

    # mssql-python
    newdict = {}
    columns_list = []
    values_list = []
    valstr = ""
    
    for start, end in _my_functions.generate_subranges(df_from_excel.shape[0], chunk_size=1000):
        # print(f"Subrange from {start} to {end}")
        for i in range(start, end):
            df_row_list = list(df_from_excel.row(i))
            for y in range(0, len(df_row_list)):
                newdict.setdefault(df_from_excel.columns[y], df_row_list[y])
            for k, v in newdict.items():
                if isinstance(v, str):
                    v = "N" + "\'" + v + "\'"
                    newdict[k] = v
                if v is None:
                    v = "NULL"
                    newdict[k] = v
                if isinstance(v, datetime.date):
                    v = "\'" + str(v) + "\'"
                    newdict[k] = v
                columns_list.append("["+str(k)+"]")
                values_list.append(str(v))
            # pprint.pprint(newdict)
            # print(values_list)
            for a, b in zip(columns_list, values_list):
                columns_str = ", ".join(columns_list)
                values_str_temp = ", ".join(values_list)
            # print(columns_str)
            # print(values_str_temp)
            valstr += "\n   (" + values_str_temp + "),"
            newdict = {}
            columns_list = []
            values_list = []
        query = "INSERT INTO dbo.BONUSDogBase_python (" + columns_str + ")\nVALUES" + valstr
        query = query[:-1] + ";"
        # print(query)
        valstr = ""
        
        with mssql_python.connect(mssql_python_uri_1) as conn:
            with conn.cursor() as cur:
                cur.execute(query)
                conn.commit()
                print("\nRows inserted:", cur.rowcount)
            print("Загружены договоры с затертыми цепочками и новые договоры")

# sys.exit()

# ------------------------------------------------------------------ обработка 126 отчета
df_126 = pl.read_excel(
    # engine="openpyxl",
    schema_overrides={
        "Дата договора лизинга": pl.Date,
        "Дата статуса": pl.Date,
        "Дата расторжения": pl.Date,
        "Дата изъятия": pl.Date
        },
    source=filename2,
    sheet_name="126",
    has_header=True
    # columns="A:B"
)
df_126 = df_126.drop("Статус ОС")
df_126 = df_126.filter(pl.col("Дата расторжения") <= cutoff_termination_date)
df_126 = df_126.sort(["Дата расторжения"], descending=True)
print(df_126.head())

# df2tables.render(df_from_excel)
# _my_functions.view_itables_html(df=df_from_excel)

# ------------------------------------------------------------------ загрузка 12 предыдущих периодов из базы договоров
query_p1 = "SELECT * FROM dbo.BONUSDogBase_python WHERE dogovor_period in ("
query_p2 = previous_12_periods
query_p3 = ")"
query = query_p1 + query_p2 + query_p3
# print(query)
df_dogovory_wide = pl.read_database_uri(
    uri=polars_uri_1,
    query=query,
    engine="connectorx",
    schema_overrides={
        "Год": pl.Int64,
        },
    )
df_dogovory_wide = df_dogovory_wide.with_columns(
    pl.when(pl.col("Основной менеджер продаж") == pl.col("Д регионы"))
    .then(pl.lit("личные продажи ДРП"))
    .otherwise(pl.lit("продажи менеджеров"))
    .alias("Кто продал")
)
df_dogovory_wide = df_dogovory_wide.sort(["LoadDt"], descending=True)
# print(df_dogovory_wide.columns)
df_dogovory_wide = df_dogovory_wide.filter(pl.col("Основной менеджер продаж (Лизинговая сделка) ").is_not_null())
# _my_functions.view_itables_html(df=df_dogovory_wide)
# print(df_dogovory_wide.head())
# print(df_dogovory_wide.columns)

# df_dogovory_wide.write_excel(filename3)
# sys.exit()

# ------------------------------------------------------------------ wide to long
df_dogovory_long = df_dogovory_wide.unpivot(
    index=["Номер договора", "Кто продал"], 
    on=['Основной менеджер продаж', 'ГМ', 'РГ', 'НО', 'ДД Москва/УСС', 'Д регионы', 'ТД', 'РД', 'НУ'], 
    variable_name="Должность", 
    value_name="ФИО"
)
df_dogovory_long = df_dogovory_long.sort(["Номер договора"], descending=False)
df_dogovory_long = df_dogovory_long.with_columns(pl.col(["ФИО"]).str.to_lowercase().alias("ФИО_join"))
# _my_functions.view_itables_html(df=df_dogovory_long)
# sys.exit()

# ------------------------------------------------------------------ загрузка изменений ФИО
df_new_fio = pl.read_excel(
    # engine="openpyxl",
    read_options={"header_row": 3},
    schema_overrides={
        "Дата изменения": pl.Date,
        },
    source=filename4,
    sheet_name="Лист_1",
    has_header=True
    # columns="A:B"
)
df_new_fio = df_new_fio.select("Дата изменения", "ФИО до изменения", "Сотрудник")
df_new_fio = df_new_fio.sort(["Дата изменения"], descending=True)
df_new_fio = df_new_fio.with_columns(pl.col(["ФИО до изменения"]).str.to_lowercase())
print(df_new_fio.head())

# sys.exit()

# ------------------------------------------------------------------ загрузка из базы demography
query = "SELECT * FROM [dbo].[Demography_report_hist] WHERE Repdt > " + list_of_12_periods[-1] + " AND Repdt < " + "\'" + str(bonus_period) + "\'"
# print(query)
df_demography = pl.read_database_uri(
    uri=polars_uri_1,
    query=query,
    engine="connectorx",
    schema_overrides={
        "Год": pl.Int64,
        },
    )
# df_demography = df_demography.with_columns((pl.col("Last_name_local") + " " + pl.col("First_name_local") + " " + pl.col("Middle_name_local")).alias("ФИО полное"))

df_demography = df_demography.with_columns(
    pl.when(pl.col("Middle_name_local").is_null())
    .then(pl.col("Last_name_local") + " " + pl.col("First_name_local"))
    .otherwise(pl.col("Last_name_local") + " " + pl.col("First_name_local") + " " + pl.col("Middle_name_local"))
    .alias("ФИО полное")
)
df_demography = df_demography.unique(subset=["ФИО полное"])
df_demography = df_demography.with_columns(pl.col(["ФИО полное"]).str.to_lowercase())

# print(df_demography.head())
# _my_functions.view_itables_html(df=df_demography)

# ------------------------------------------------------------------ добавление статуса договора, СНИЛС, нового ФИО
df_dogovory_long = df_dogovory_long.join(df_126, left_on="Номер договора", right_on="Договор лизинга", how="left")
df_dogovory_long = df_dogovory_long.join(df_new_fio, left_on="ФИО_join", right_on="ФИО до изменения", how="left")
df_dogovory_long = df_dogovory_long.join(df_demography, left_on="ФИО_join", right_on="ФИО полное", how="left")
df_dogovory_long = df_dogovory_long.select("Номер договора", "Кто продал", "Статус договора лизинга", "Должность", "ФИО", "Сотрудник", "Social_number")
df_dogovory_long = df_dogovory_long.rename({
    "Сотрудник": "ФИО_новое",
    })
df_dogovory_long = df_dogovory_long.with_columns(
        pl.when(pl.col("ФИО_новое").is_null())
        .then(pl.col("ФИО"))
        .otherwise(pl.col("ФИО_новое"))
        .alias("ФИО_новое")
    )
# df_dogovory_long = df_dogovory_long.filter(pl.col("ФИО").is_not_null() & pl.col("Social_number").is_null())

# df2tables.render(df_dogovory_long)
# _my_functions.view_itables_html(df=df_dogovory_long)
# df_dogovory_long.write_excel(filename3)
# sys.exit()

# ------------------------------------------------------------------ расчет коэфф расторжений
df_koeff = pl.DataFrame([])

for i in ["продажи менеджеров", "личные продажи ДРП"]:
    _my_functions.print_line("hyphens")
    print(i + "\n")

    # df_vsego = df_dogovory_long.filter(pl.col("Кто продал") == i).group_by(["Social_number"]).agg(
    df_vsego = df_dogovory_long.filter(pl.col("ФИО").is_not_null() & (pl.col("Кто продал") == i)).group_by(["Social_number"]).agg(
        pl.col("Номер договора").count().alias("всего договоров")
    )
    # print(df_vsego)
    # print("Всего договоров " + str(df_vsego["всего договоров"].sum()) + "\n")

    df_rastorgn = df_dogovory_long.filter((pl.col("Кто продал") == i) & (pl.col("Статус договора лизинга") == "Расторгнут") ).group_by(["Social_number"]).agg(
        pl.col("Номер договора").count().alias("расторгнутых договоров")
    )
    # print(df_rastorgn)
    # print("Расторгнутых договоров " + str(df_rastorgn["расторгнутых договоров"].sum()) + "\n")

    df_result = df_vsego.join(df_rastorgn, on="Social_number", how="left")
    df_result = df_result.with_columns(
        pl.when(pl.col("расторгнутых договоров").is_null())
        .then(pl.lit(0))
        .otherwise(pl.col("расторгнутых договоров"))
        .alias("расторгнутых договоров")
    )
    if i == "продажи менеджеров":
        df_result = df_result.with_columns(
            (1-pl.col("расторгнутых договоров") / pl.col("всего договоров")).alias("коэфф_р")
        )
    if i == "личные продажи ДРП":
        df_result = df_result.with_columns(
            (1-pl.col("расторгнутых договоров") / pl.col("всего договоров")).alias("коэфф_р_лДРП")
        )
        # df_result.write_excel(filename3)
        # sys.exit()
    print("df_result")
    print(df_result)
    # print("Всего договоров " + str(df_result["всего договоров"].sum()))
    # print("Расторгнутых договоров " + str(df_result["расторгнутых договоров"].sum()) + "\n")

    if i == "продажи менеджеров":
        df_koeff = pl.concat([df_koeff, df_result], how="diagonal")
    if i == "личные продажи ДРП":
        # df_koeff = df_koeff.join(df_result, on="Social_number", how="full")
        df_koeff = df_koeff.join(df_result, on="Social_number", how="left")
    # sys.exit()

_my_functions.print_line("hyphens")
df_koeff = df_koeff.rename({
    "всего договоров_right": "всего договоров_лДРП",
    "расторгнутых договоров_right": "расторгнутых договоров_лДРП"
    })
print(df_koeff)

# df_koeff.write_excel(filename3)
# sys.exit()

# ------------------------------------------------------------------ добавление коэффициента расторжений
df_dogovory_long = df_dogovory_long.join(df_koeff, on="Social_number", how="left")
# df_dogovory_long = df_dogovory_long.filter(pl.col("ФИО").is_not_null() & pl.col("Social_number").is_null())
# df_dogovory_long = df_dogovory_long.filter(pl.col("ФИО").is_not_null() & pl.col("коэфф_р").is_null())
df_dogovory_long = df_dogovory_long.filter(pl.col("ФИО").is_not_null())
df_dogovory_long = df_dogovory_long.sort(["Social_number"])

# df2tables.render(df_dogovory_long)
_my_functions.view_itables_html(df=df_dogovory_long)

df_dogovory_long.write_excel(filename3)

# ------------------------------------------------------------------ отбор работников, по которым менялось ФИО
"""
_my_functions.print_line("hyphens")
print("Работники с изменениями ФИО:")
print("\n")
sn_list = df_dogovory_long["Social_number"].to_list()
fio_list = df_dogovory_long["ФИО"].to_list()
sn_fio_dict = {}
for a,b in zip(sn_list, fio_list):
    sn_fio_dict.setdefault(a, [])
    if b not in sn_fio_dict[a]:
        sn_fio_dict[a].append(b)
for k,v in sn_fio_dict.items():
    if len(v) > 1:
        print(k)
        print(v)
        print("\n")
"""

# sys.exit()