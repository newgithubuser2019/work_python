# IMPORTS
import datetime
import decimal
import json
import os
import pprint
import re
import shutil
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
from pandas.tseries.offsets import DateOffset
from sqlalchemy import create_engine

import _my_functions

pd.set_option("display.max_rows", 1500)
pd.set_option("display.max_columns", 100)
pd.set_option("max_colwidth", 30)
pd.set_option("expand_frame_repr", False)
# ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# global variables
USERPROFILE = os.environ["USERPROFILE"]

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# empty lists

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------
# empty dictionaries

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------
# empty dataframes
# df = pd.DataFrame()

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# default lists

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------
# default dictionaries

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# prompts for user input
# prompt1 = "\nРеализация: "

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------
# user inputs
# inp1 = input(prompt1)

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------
# file paths
# path_1 = USERPROFILE + "\\Documents\\Работа\\отчетность\\ежедневно\\накопительный отчет\\время поднятия кормушки\\"
# listoffiles_1 = os.listdir(path_1)

# file names
# filename0 = USERPROFILE + "\\Documents\\Работа\\отчетность\\ежедневно\\накопительный отчет\\_промежуточный файл df_впк.xlsx"
filename1 = "P:\\Documents\\ДБ\\СРП\\Компенсации и льготы\\Премирование\\!СМОТы, встречи, материалы\\!Аналитика\\своды по премиям\\ремаркетинг свод.xlsx"
filename2 = "P:\\Documents\\ДБ\\СРП\\Компенсации и льготы\\Потапов Д\\отчеты\\контакт\\126 отчет\\2026\\06\\126.xlsx"

# sql_username = getpass.getuser()
# sql_password = getpass.getpass(prompt="SQL password: ", stream=None)
# sql_password = getpass.getpass(prompt="SQL password: ", stream=None, echo_char="*") # echo_char parameter added in python 3.14
# uri1 = "mssql://" + sql_username + ":" + sql_password + "@vls-sql-zup-dev:1433/HR_CAB"
# print(uri1)
uri1 = "mssql://vls-sql-zup-dev:1433/HR_CAB?trusted_connection=true"

# ------------------------------------------------------------------------------------------------------------------------------------------
"""
df_from_excel = pd.read_excel(
    filename1,
    sheet_name="продажи",
    # index_col=0,
    # engine = "openpyxl",
    header=0,
    usecols = "A:AB",
    # usecols = "A,C,G,H,I,K",
    )
# print(df_from_excel.info())
"""
#df2tables.render(df_from_excel)

# ------------------------------------------------------------------------------------------------------------------------------------------

# conn = ""
# query = "SELECT * FROM dbo.Termination_URPA"
# df_arrow = cx.read_sql(conn, query, protocol='text', return_type='arrow')
# df = df_arrow.to_pandas()

prompt1 = "\nгод отчетного периода: "
prompt2 = "\nмесяц отчетного периода: "
prompt3 = "\nДля продолжения необходимо загрузить договоры с затертыми цепочками и новые договоры. Выполнена ли загузка? "

inp1 = input(prompt1)
inp2 = input(prompt2)

# ------------------------------------------------------------------ даты и периоды
bonus_period = datetime.date(int(inp1), int(inp2), 1)
# print(bonus_period.month)
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
    uri=uri1,
    query=query,
    engine="connectorx",
    # schema_overrides={
        #"Год": pl.Int64,
        #},
    )
print(df_from_sql.shape[0])

if df_from_sql.shape[0] == 0:
    # print("Не загружены новые договоры / договоры с затертыми цепочками")

    if len(str(inp2)) == 2:
        month = str(inp2)
    else:
        month = "0" + str(inp2)
    filename = "P:\\Documents\\ДБ\\СРП\\Компенсации и льготы\\Потапов Д\\коэффициент расторжений\\" + str(inp1) + "\\" + month + "\\договоры\\договоры.xlsx"
    # print(filename)

    df_from_excel = pl.read_excel(
        # engine="openpyxl",
        schema_overrides={
            "Дата удержания ": pl.Date,
            "LoadDt": pl.Date,
            "dogovor_period": pl.Date,
            },
        source=filename,
        sheet_name="Лист1",
        has_header=True
        # columns="A:B"
    )
    # print(df_from_excel.head())
    df_rastorgnut = df_from_excel.filter((pl.col("Комментарий") == "расторжение"))
    df_new = df_from_excel.filter((pl.col("Комментарий") != "расторжение") | pl.col("Комментарий").is_null())
    # print(df_rastorgnut)

    # удаление расторгнутых договоров из базы
    rastorgn_dogovory = df_rastorgnut["Номер договора"].to_list()
    # print(rastorgn_dogovory)
    rastorgn_dogovory_query = []
    for i in rastorgn_dogovory:
        rastorgn_dogovory_query.append("N" + "\'" + i + "\'")
    print(rastorgn_dogovory_query)

    query_p1 = "DELETE FROM dbo.BONUSDogBase_python WHERE [Номер договора] in ("
    query_p2 = ", ".join(rastorgn_dogovory_query)
    query_p3 = ")"
    query = query_p1 + query_p2 + query_p3
    print(query)
    # query = "SELECT * FROM dbo.BONUSDogBase_python"

    # apache arrow adbc
    """
    from adbc_driver_manager import dbapi
    with dbapi.connect(uri1) as conn:
        with conn.cursor() as cur:
            cur.execute("query")
            conn.commit()
            print("Rows deleted:", cur.rowcount)
    """

    # mssql-python
    conn_str = "Server=vls-sql-zup-dev,1433;Database=HR_CAB;Trusted_connection=yes;TrustServerCertificate=yes"
    # conn = mssql_python.connect(conn_str)
    
    with mssql_python.connect(conn_str) as conn:
        with conn.cursor() as cur:
            cur.execute(query)
            conn.commit()
            print("Rows deleted:", cur.rowcount)
    print("Удалены договоры, по которым затерты цепочки")

    # загрузка новых договоров в базу
    newdict = {}
    columns_list = []
    values_list = []
    for i in range(0, df_from_excel.shape[0]):
    # for i in range(0, 1):
        print(i)
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
            values_str = ", ".join(values_list)
        # print(columns_str)
        # print(values_str)
        query = "INSERT INTO dbo.BONUSDogBase_python (" + columns_str + ")\nVALUES (" + values_str + ")"
        # print(query)
        newdict = {}
        columns_list = []
        values_list = []

    sys.exit()
    inp3 = input(prompt3)
    if inp3.lower() == "да":
        pass
    else:
        print("проверка сработала")
        sys.exit()
    # print(df_from_excel.shape[0])
    # print(df_from_excel)

    # apache arrow adbc
    """
    df_from_excel.write_database(
                    table_name="BONUSDogBase_python",
                    connection=uri1,
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
    sys.exit()

# ------------------------------------------------------------------ обработка 126 отчета
df_from_excel = pl.read_excel(
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
df_from_excel = df_from_excel.drop("Статус ОС")
df_from_excel = df_from_excel.filter(pl.col("Дата расторжения") <= cutoff_termination_date)
df_from_excel = df_from_excel.sort(["Дата расторжения"], descending=True)
print(df_from_excel.head())

_my_functions.view_itables_html(df=df_from_excel)

# ------------------------------------------------------------------ загрузка 12 предыдущих периодов из базы договоров
query_p1 = "SELECT * FROM dbo.BONUSDogBase_python WHERE dogovor_period in ("
query_p2 = previous_12_periods
query_p3 = ")"
query = query_p1 + query_p2 + query_p3
# print(query)
df_from_sql = pl.read_database_uri(
    uri=uri1,
    query=query,
    engine="connectorx",
    schema_overrides={
        "Год": pl.Int64,
        },
    )
df_from_sql = df_from_sql.sort(["LoadDt"], descending=True)
print(df_from_sql.head())
# df2tables.render(df_from_sql)

sys.exit()
