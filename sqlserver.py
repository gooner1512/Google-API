import pandas as pd
import pyodbc

class Database:
    def __init__(self):
        self.conn = self.init_connection()

    @staticmethod
    def init_connection():
        return pyodbc.connect(
            "DRIVER={ODBC Driver 17 for SQL Server};SERVER="
            + ''
            + ";DATABASE="
            + 'BASE'
            + ";UID="
            + 'UserReport'
            + ";PWD="
            + ''
        )    
      
    def run_query(_self, query):
        with _self.conn.cursor() as cur:
            cur.execute(query)
            # Fetch the rows
            rows = cur.fetchall()

            # Get the column names from the cursor description
            column_names = [column[0] for column in cur.description]

            # Separate the columns into separate lists
            columns = []
            for i in range(len(column_names)):
                column = [row[i] for row in rows]
                columns.append(column)

            # Create the DataFrame
            df = pd.DataFrame({column_names[i]: columns[i] for i in range(len(column_names))})
            return df
    
    def run_query_multi_tables(_self, query):
        with _self.conn.cursor() as cur:
            cur.execute(query)
            if not cur.description:
                return [] # no results
            rows = cur.fetchall()
            columns = [column[0] for column in cur.description]
            df_list = []
            df_list.append(pd.DataFrame.from_records(rows, columns=columns))
            
            while (cur.nextset()): 
                rows = cur.fetchall()
                columns = [column[0] for column in cur.description]
                df_list.append(pd.DataFrame.from_records(rows, columns=columns))

            return df_list