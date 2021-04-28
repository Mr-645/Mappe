import sqlite3

sql_table_path = "C:/Users/drift/OneDrive/PROJECTS/Mappe/testing_ground/test.db"

connection = sqlite3.connect(sql_table_path)
query = "SELECT * FROM main_table_db"
result = connection.execute(query)

for row_number, row_data in enumerate(result):
    print(row_number)
    for column_number, data in enumerate(row_data):
        print(row_number, column_number, (str(data)))

print(f"\n")

query2 = "PRAGMA table_info(main_table_db)"
result2 = connection.execute(query2)

header_list = result2.fetchall()

# print(f"\nType = {type(header_list)}")
# print(f"{header_list}")

# print(f"\nType = {type(header_list[1])}")
# print(f"{header_list[1]}")

# print(f"\nType = {type((header_list[0])[1])}")
# print(f"{(header_list[0])[1]}")

# print(f"\nType = {type((header_list[1])[1])}")
# print(f"{(header_list[1])[1]}")

header_list_2 = []

for i in header_list:
    header_list_2.append(i[2])

print(f"\n{header_list_2}\n")

# =================================================

connection.execute("SET DATEFORMAT dmy")
query3 = "PRAGMA table_info(main_table_db)"
result3 = connection.execute(query3)

header_list_3 = result3.fetchall()

header_list_4 = []

for i in header_list:
    header_list_4.append(i[2])

print(f"\n{header_list_4}\n")

# ==================================================

connection.close()