from tabula import convert_into
table_file = r"input.pdf"
output_csv = r"output.csv"
df = convert_into(table_file, output_csv, output_format='csv', lattice=True, stream=False, pages="all")