import sqlite3

conn = sqlite3.connect('testing_ground/test.db')

print("Opened database successfully")

conn.execute('''
CREATE TABLE IF NOT EXISTS [main_table_db] (
[ID] VARCHAR NOT NULL,
[doc_name] VARCHAR NOT NULL,
[fac_location] VARCHAR NOT NULL,
[mod_date] DATE NOT NULL,
[expiry_status] BIT NOT NULL,
[file_uri] VARCHAR NOT NULL
);
''')
print("Table created successfully")

conn.execute("INSERT INTO main_table_db VALUES ('S-452-001','SOP_Moulder_2_task_1','452 - Line 2','2021/03/16','0','Documents/SOP_Moulder_2_task_1.docx')")
conn.execute("INSERT INTO main_table_db VALUES ('S-452-002','SOP_Moulder_2_task_2','452 - Line 2','2021/03/17','0','Documents/SOP_Moulder_2_task_2.docx')")

conn.commit()
print("Records created successfully")

conn.close()