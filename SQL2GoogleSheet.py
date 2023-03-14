from sqlserver import Database
import pandas as pd

db = Database()

query="""
EXEC USP_KIDS_TEST '20230201','20230202','1'
"""
dataframes = db.run_query_multi_tables(query)
df_0 = dataframes[0]
df_1 = dataframes[1]
df_2 = dataframes[2]
df_3 = dataframes[3]
df_4 = dataframes[4]

recordset = df_0.values.tolist()