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

import numpy as np
import openpyxl
import pandas as pd
import plotly.express as px
import connectorx as cx
import polars as pl
import getpass

import df2tables
from pandas.tseries.offsets import DateOffset
# from itables.streamlit import interactive_table

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

# sql_username = getpass.getuser()
# sql_password = getpass.getpass(prompt="SQL password: ", stream=None)
# sql_password = getpass.getpass(prompt="SQL password: ", stream=None, echo_char="*") # echo_char parameter added in python 3.14
# uri1 = "mssql://" + sql_username + ":" + sql_password + "@vls-sql-zup-dev:1433/HR_CAB"
# print(uri1)
# #sys.exit()

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

# conn = ""
# query = "SELECT * FROM dbo.Termination_URPA"
# df_arrow = cx.read_sql(conn, query, protocol='text', return_type='arrow')
# df = df_arrow.to_pandas()

# sys.exit()

query = "SELECT * FROM dbo.Termination_URPA"
df_from_sql = pl.read_database_uri(
    uri=uri1,
    query=query,
    engine="connectorx"
    )
print(df_from_sql.head())
# df2tables.render(df_from_sql)

sys.exit()
