import tabula
import pandas as pd

# Define the paths
pdf_file = "sample.pdf"
output_excel_file = "auto.xlsx"

# Read PDF and convert to Excel
tables = tabula.read_pdf(pdf_file, pages="all", multiple_tables=True)

# Combine all tables into a single DataFrame (if multiple tables are present)
df = pd.concat(tables, ignore_index=True)

# Save the DataFrame to Excel
df.to_excel(output_excel_file, index=False)

print(f"PDF '{pdf_file}' converted to Excel '{output_excel_file}' successfully.")

