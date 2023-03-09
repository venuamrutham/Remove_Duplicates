import pandas as pd
import tkinter as tk
from tkinter import filedialog
import webbrowser
# Create a tkinter window
root = tk.Tk()
root.withdraw() # hide the root window


# Prompt the user to select the input file using a file picker dialog
file_path = filedialog.askopenfilename()

# Read the Excel file
df = pd.read_excel(file_path)

# Drop duplicates
df.drop_duplicates(inplace=True)

# Save cleaned data to a new Excel file
df.to_excel('cleaned_file.xlsx', index=False)

# Generate the HTML table from the cleaned data
table_html = df.to_html(index=False)

# Write the HTML file to display the table
with open('index.html', 'w') as f:
    html_template = """
    <!DOCTYPE html>
    <html>
    <head>
        <title>Cleaned Data</title>
    </head>
    <body>
        <h1>Cleaned Data</h1>
        {table_html}
    </body>
    </html>
    """
    f.write(html_template.format(table_html=table_html))
    print('finished cleaning up, opening the file')
    webbrowser.open_new_tab('index.html')
    print('successfully opened')

