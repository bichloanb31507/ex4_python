from ast import Constant
import traceback
import pandas as pd

def getData():
        try:
            data = []
            df = pd.read_excel(r'D:\input.xlsx', sheet_name='MAU', usecols='A:H', skiprows=10, nrows=52)
            for row in df.iterrows():
                row_data = []
                for value in row[1]:
                    row_data.append(value)
                data.append(row_data)
                print("getData thanh cong")
            return data;
        except:
            traceback.print_exc()
            print("getData that bai")
            
def importData(conn, data):
        try:
            cursor = conn.cursor()
            for row_data in data:
                values = (row_data[0],row_data[1],row_data[2],row_data[3],row_data[4],row_data[5],row_data[6],row_data[7])
                cursor.execute(Constant.ADD_SQL, values)    
                conn.commit()
            print("importData thanh cong")
        except:
            if conn is not None:
                conn.rollback()
                conn.close()
            traceback.print_exc()
            print("importData that bai")