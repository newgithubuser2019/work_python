# IMPORTS
import datetime
import decimal
import json
# import math
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
# from sqlalchemy import create_engine

import _my_functions

pd.set_option("display.max_rows", 1500)
pd.set_option("display.max_columns", 100)
pd.set_option("max_colwidth", 30)
pd.set_option("expand_frame_repr", False)

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------ global variables
USERPROFILE = os.environ["USERPROFILE"]

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------ prompts for user input
# prompt1 = "\nРеализация: "

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------ user inputs
# inp1 = input(prompt1)

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------ file paths
# path_1 = USERPROFILE + "\\Documents\\Работа\\отчетность\\ежедневно\\накопительный отчет\\"
# listoffiles_1 = os.listdir(path_1)

# ------------------------------------------------------------------ file names
filename1 = "P:\\Documents\\ДБ\\СРП\\Компенсации и льготы\\Премирование\\!СМОТы, встречи, материалы\\!Аналитика\\своды по премиям\\ремаркетинг свод.xlsx"
filename2 = "P:\\Documents\\ДБ\\СРП\\Компенсации и льготы\\Потапов Д\\отчеты\\контакт\\126 отчет\\2026\\06\\126.xlsx"

# ------------------------------------------------------------------ database URIs
# sql_username = getpass.getuser()
# sql_password = getpass.getpass(prompt="SQL password: ", stream=None)
# sql_password = getpass.getpass(prompt="SQL password: ", stream=None, echo_char="*") # echo_char parameter added in python 3.14
# uri1 = "mssql://" + sql_username + ":" + sql_password + "@vls-sql-zup-dev:1433/HR_CAB"
# print(uri1)
uri1 = "mssql://vls-sql-zup-dev:1433/HR_CAB?trusted_connection=true"

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------ load dataframe from excel
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

"""
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
print(df_from_excel.head())
"""

# df2tables.render(df_from_excel)
# _my_functions.view_itables_html(df=df_from_excel)

# ------------------------------------------------------------------ load dataframe from sql

# conn = ""
# query = "SELECT * FROM dbo.Termination_URPA"
# df_arrow = cx.read_sql(conn, query, protocol='text', return_type='arrow')
# df = df_arrow.to_pandas()

"""
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
"""
# df2tables.render(df_from_sql)
# _my_functions.view_itables_html(df=df_from_sql)

# sys.exit()

# ---------------------------------------------------------------------------------------------------------------------------------------------------------------
