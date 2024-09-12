import pandas as pd
from docx import Document
from tabulate import tabulate
from collections import defaultdict
import tkinter as tk
from tkinter import filedialog
import webbrowser
import os


def extract_table_data(doc_path):
    doc = Document(doc_path)
    table_data = []
    for table in doc.tables:
        for row in table.rows:
            row_data = [cell.text.strip() for cell in row.cells]
            table_data.append(row_data)
    return table_data, doc


def display_columns_and_choose(data):
    # Create DataFrame
    df = pd.DataFrame(data)

    # Display first row as column names with indices
    first_row = df.iloc[0]
    col_indices = list(range(len(first_row)))
    print("\nColumn Names and Indices:")
    for idx, col_name in zip(col_indices, first_row):
        print(f"{idx} - {col_name}")

    # Find the column with "BARWA" in its name and the next two columns
    barwa_col_idx = next((i for i, col_name in enumerate(first_row) if "BARWA" in col_name.upper()), None)
    if barwa_col_idx is not None:
        default_cols = [barwa_col_idx, barwa_col_idx + 1, barwa_col_idx + 2]
    else:
        default_cols = []

    print("\nDefault columns chosen based on 'BARWA' criteria:", default_cols)

    # Ask user if they want to choose other columns
    user_input = input("Do you want to choose other columns? (yes/no): ").strip().lower()
    if user_input == 'yes':
        chosen_indices = input("Enter column indices (e.g., 1-5 for columns 1,2,3,4,5): ").strip()
        chosen_indices = [int(i) for i in chosen_indices.split('-')]
    else:
        chosen_indices = default_cols

    return df, chosen_indices


def process_column_data(column):
    # Remove empty elements and leading/trailing spaces, replace "–" with "-", split by newline
    processed_col = [item.strip().replace('–', '-').replace(' ', '') for sublist in column for item in
                     sublist.split('\n') if item.strip()]

    # Partition into number and type
    partitioned_data = defaultdict(int)
    for item in processed_col:
        if '-' in item:
            num, typ = item.split('-', 1)
            if num.isdigit():  # Check if num is a valid integer
                partitioned_data[typ] += int(num)
            else:
                print(f"Skipping invalid number: {num} in item: {item}")

    return partitioned_data


def get_summed_data(df, chosen_indices):
    summed_data = []
    for idx in chosen_indices:
        chosen_column = df.iloc[:, idx].tolist()
        processed_data = process_column_data(chosen_column)
        summed_data.append(processed_data)
    return summed_data


def generate_and_open_html(summed_data, original_column_names):
    html_output = '<html><head><style>'
    html_output += 'table { width: 100%; border-collapse: collapse; }'
    html_output += 'th, td { border: 1px solid black; padding: 8px; text-align: left; }'
    html_output += 'th { background-color: #f2f2f2; }'
    html_output += 'tr:nth-child(even) { background-color: #f9f9f9; }'
    html_output += '</style></head><body>'

    for i, column_data in enumerate(summed_data):
        table_name = original_column_names[chosen_indices[i]]
        headers = ['Type', 'Quantity']

        # Sort data by quantity (value) in descending order
        sorted_data = sorted(column_data.items(), key=lambda x: x[1], reverse=True)

        html_output += f"<h2>Table {i + 1}: {table_name}</h2>\n"
        html_output += '<table>\n'
        html_output += '<tr>' + ''.join([f'<th>{header}</th>' for header in headers]) + '</tr>\n'
        for typ, num in sorted_data:
            html_output += f'<tr><td>{typ}</td><td>{num}</td></tr>\n'
        html_output += '</table><br>\n'

    html_output += '</body></html>'

    # Save HTML output to a temporary file
    temp_file = 'summed_data_output.html'
    with open(temp_file, 'w', encoding='utf-8') as file:
        file.write(html_output)

    # Open the HTML file in the default web browser
    webbrowser.open(f'file://{os.path.abspath(temp_file)}')


def choose_file():
    # Create a file dialog for choosing the DOCX file
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
    return file_path


# Main code
doc_path = choose_file()
if not doc_path:
    print("No file selected.")
else:
    table_data, doc = extract_table_data(doc_path)

    # Display extracted table data
    print("\nExtracted Table Data:")
    print(tabulate(table_data))

    # Process and choose columns
    df, chosen_indices = display_columns_and_choose(table_data)

    # Get original column names for output
    original_column_names = df.columns.tolist()

    # Get summed data for validation
    summed_data = get_summed_data(df, chosen_indices)

    # Generate HTML and open it in the default web browser
    generate_and_open_html(summed_data, original_column_names)
