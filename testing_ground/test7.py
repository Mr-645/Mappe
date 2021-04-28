import pandas as pd
from docx.api import Document

document_1 = Document("C:/Users/drift/OneDrive/PROJECTS/Mappe/testing_ground/SOP_docs/test SOP AMF CIP.docx")

table = document_1.tables[0]

data = []

keys = None
for i, row in enumerate(table.rows):
    text = (cell.text for cell in row.cells)

    if i == 0:
        keys = tuple(text)
        continue
    row_data = dict(zip(keys, text))
    data.append(row_data)
    print(data)

df = pd.DataFrame(data)
print(df)