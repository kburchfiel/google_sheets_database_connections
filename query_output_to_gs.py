# Query Output to Google Slides
# Kenneth Burchfiel
# Program is released under the MIT License

'''This program shows how to upload the results of database queries to a Google
Sheets File. The program uses a sample SQLite database containing fictional
test score data; however, you can also connect to an online database using
SQLalchemy. See my Python Database Utilities repository (available at 
https://github.com/kburchfiel/python_database_utilities) for examples.'''

'''More documentation will be provided in the future. I will probably also
convert the .py file to an .ipynb file for easier readability.'''

import sqlite3
import pandas as pd
import getpass
import gspread
from gspread_dataframe import set_with_dataframe
import time

con = sqlite3.connect('test_scores.db') # I initialized 'test.db' simply be 
# creating an empty file in my folder and giving it that name. 

df_scores = pd.read_excel('scores_by_program_enrollment.xlsx') # This idea for 
# importing a spreadsheet into a DataBase came from Stack Overflow user 
# Tennessee Leeuwenburg (see https://stackoverflow.com/a/28802613/13097194).

print(df_scores)

df_scores.to_sql('Scores', con = con, if_exists = 'replace') # See 
# https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.to_sql.html

cur = con.cursor()

gc = gspread.service_account(pd.read_csv(
    '..\\key_paths\\key_paths.csv').iloc[0,1])

query_list = []

query_list.append("Select * from Scores limit 5")
query_list.append("Select Student_ID, School, Grade from Scores limit 50")

query_dict_list = []
for i in range(len(query_list)):
    query_dict_list.append(
        {"query_id":"Query_"+str(i),"query_text":query_list[i]})

results_workbook = gc.open_by_key(
    '1jPPz4YW5v5repoJXpXXJ3VrivK21lv1VYLvQIvTEyxE')

df_query_index = pd.DataFrame(query_dict_list)

print(df_query_index)

query_index_sheet = results_workbook.get_worksheet(0)
query_index_sheet.clear()
query_index_sheet_title = 'query_index'
query_index_sheet.update_title(query_index_sheet_title)
set_with_dataframe(query_index_sheet, df_query_index, include_index = True)



for i in range(len(query_list)):
    start_time = time.time()
    print("Now on Query",i)
    df_query = pd.read_sql(sql = query_list[i], con = con) # This was a method 
    # I had originally learned about when converting database concent accessed 
    # through pyodbc to Pandas DataFrames. It works with sqlite3 databases also.
    # print(df_query) # Helpful for debugging
    query_sheet = results_workbook.get_worksheet(i+1) # A +1 offset is used 
    # because sheet 0 contains the index list.
    query_sheet.clear()
    query_sheet_title = 'Query_'+str(i)
    query_sheet.update_title(query_sheet_title)
    set_with_dataframe(query_sheet, df_query, include_index = True)
    end_time = time.time()
    length = end_time - start_time
    print("Time operation took (in seconds):",'{:.3f}'.format(length))

