from docx import Document
from tabulate import tabulate

wordDoc = Document(r'C:\Users\Paulo\Desktop\izm.docx')
tableData = []

for table in wordDoc.tables:
    for row in table.rows:
        rowData = []
        for cell in row.cells:
            rowData.append(cell.text.strip())
        tableData.append(rowData)

# Save data to a list
extracted_data = tableData

# Display table
print(tabulate(extracted_data))