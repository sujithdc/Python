import pandas as pd


current_row = 0
sheet_num = 1
input_total = 0
output_total = 0

# path to the file you want to extract data from
src = r'd:\test1\attachments\OCS 5.5 test numbers.xlsx'

xl_file = pd.ExcelFile(src)

dfs = {sheet_name: xl_file.parse(sheet_name) 
          for sheet_name in xl_file.sheet_names}
print (dfs)
