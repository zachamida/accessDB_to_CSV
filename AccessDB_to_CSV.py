# This script perform the following:
# 1) remove commas from accessDB
# 2) Export tables in accessDB to .csv files
# File created by: Zachary Hamida (2023)

# It is required to install 
# Microsoft Access Database Engine 2016 Redistributable
# https://www.microsoft.com/en-US/download/details.aspx?id=54920

import pyodbc
import csv

def acc2csv(mdb_file):
    try:
        connection_string = f"Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};Dbq={mdb_file};"
        connection = pyodbc.connect(connection_string, autocommit=True)

        tables = []
        for row in connection.cursor().tables():
            if row.table_type == "TABLE":
                col_query = f"SELECT * FROM [{row.table_name}] WHERE 1 = 0;"  # Retrieve no rows, only the structure
                with connection.cursor() as cursor:
                    cursor.execute(col_query)
                for column in cursor.description:
                    replace_query = f"UPDATE [{row.table_name}] SET [{column[0]}] = REPLACE([{column[0]}], ',', ' ') WHERE INSTR([{column[0]}], ',') > 0;"
                    with connection.cursor() as cursor_col:
                        cursor_col.execute(replace_query)
                # export to .csv file
                export_table_to_csv(connection, row.table_name, f"{row.table_name}.csv")
        
        connection.close()
        print("Column update completed successfully.")
        return tables
    except Exception as e:
        print("Error:", e)

def export_table_to_csv(connection, table_name, csv_file):
    try:
        select_query = f"SELECT * FROM [{table_name}];"
        
        with connection.cursor() as cursor:
            cursor.execute(select_query)
            rows = cursor.fetchall()
        
        with open(csv_file, 'w', newline='', encoding='utf-8') as csvfile:
            csv_writer = csv.writer(csvfile)
            
            # Write the header (column names)
            column_names = [column[0] for column in cursor.description]
            csv_writer.writerow(column_names)
            
            # Write the data rows
            for row in rows:
                csv_writer.writerow(row)
        
        print(f"Table '{table_name}' exported to '{csv_file}' successfully.")
    except Exception as e:
        print("Error:", e)

# Set the path to your MS access file
mdb_file_path = "C:/Users/username/Desktop/file.accdb"

# Call the function to perform the column update
acc2csv(mdb_file_path)