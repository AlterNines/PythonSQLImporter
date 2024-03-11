import pyodbc as dbc
from sqlalchemy import create_engine
from sqlalchemy.engine import URL
from dotenv import load_dotenv
import os

load_dotenv()


class DatabaseHandler:
    def __init__(self):
        self.driver = os.getenv('db_Driver')
        self.server_name = os.getenv('db_Server')
        self.database = os.getenv('Database')

        connection_url = URL.create('mssql+pyodbc',
                                    query={'odbc_connect': f'DRIVER={self.driver};'
                                                           f'SERVER={self.server_name};'
                                                           f'DATABASE={self.database};'
                                                           f'Trusted_Connection=yes;'})
        self.engine = create_engine(connection_url, module=dbc)

    def insert_data(self, table_name, data):
        try:
            data.to_sql(table_name, con=self.engine, if_exists='replace', index=False)
        except Exception as e:
            print(f"Error inserting data into {table_name}: {str(e)}")

