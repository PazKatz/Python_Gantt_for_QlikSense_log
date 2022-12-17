import os #read all log files with loop for
import re
# import datetime  # foramt hh:mm
import xlsxwriter as xls  # output excel file
from openpyxl.workbook import Workbook
workbook = xls.Workbook("Gant_Task_Qlik.xlsx")  # Name Excel file
worksheet = workbook.add_worksheet("Tasks_timeline")  # Name Excel TAB
worksheet.write(0,0,"Task_Name")
worksheet.write(0,1,"Start")
worksheet.write(0,2,"End")
worksheet.write(0,3,"Duration")
################################################
# os # read all log files with loop for
# re # search lines with patterns strings + readlines

folderpath = input('Please insert a folder path of log files from Qlik Sense\n')  # r"C:\Users\97250\Desktop\logs"
filepaths = [os.path.join(folderpath, name) for name in os.listdir(folderpath)]
all_files = []

# Insert start+end time from log text file of qlik sense :
index_xls_r = 1  # row index promoted
col_xls_start = 1  # column static num
col_xls_end = 2  # column static num
col_xls_duration = 3  # column static num
for path in filepaths: # building array of all log text files
    with open(path, "rt",encoding='utf-8') as my_log:
      a = [re.search(r'\d{8}[A-Z]{1}\d{6}', line)[0] for line in my_log.readlines() if 'Execution started' in line or 'Execution finished' in line]

####################################################
# Create Excel File:
    start = a[0][10:12]+':'+a[0][12:14]  # +':'+a[0][14:15]
    end = a[1][10:12]+':'+a[1][12:14]  # +':'+a[1][14:15]

    start_with_date = '01-01-1900-'+start
    end_with_date = '01-01-1900-' +end

    # print('Time Start Task:', start)
    # print('Time End Task:', end)

    import datetime
    start_time = datetime.datetime.strptime(start, '%H:%M').time() #אין צורך -אפשר למחוק
    end_time = datetime.datetime.strptime(end, '%H:%M').time() #אין צורך -אפשר למחוק

    start_with_date_format = datetime.datetime.strptime(start_with_date, '%d-%m-%Y-%H:%M')
    end_with_date_format = datetime.datetime.strptime(end_with_date, '%d-%m-%Y-%H:%M')



    #print(start_with_date_format)
    #print(start_time)
    from datetime import datetime, date  # חישוב זמן טאסק
    duration = (datetime.combine(date.today(), end_time) - datetime.combine(date.today(), start_time))

    format_hh_mm = workbook.add_format({'num_format': 'hh:mm'})  # excel cell output hh:mm
    format2 = workbook.add_format({'num_format': 'dd/mm/yy hh:mm'})  # excel cell output hh:mm
    # worksheet.write(1,1,start) #str format
    # worksheet.write(1,2,end)
    # worksheet.write(1,3,duration)

    worksheet.write(index_xls_r,col_xls_start, start_time, format_hh_mm)
    worksheet.write(index_xls_r,col_xls_end, end_time, format_hh_mm)

    # worksheet.write(index_xls_r,col_xls_start, start_with_date_format, format2)
    # worksheet.write(index_xls_r,col_xls_end, end_with_date_format, format2)

    worksheet.write(index_xls_r, col_xls_duration, duration, format_hh_mm)
    worksheet.write(index_xls_r, 0, path)

    #worksheet.write('B2', start_time, format2)
    #worksheet.write('C2', end_time, format2)
    #worksheet.write('D2', duration, format2)
    index_xls_r += 1
#################################################
# import pandas as pd
# import matplotlib.pyplot as plt
# import numpy as np
##############################
from pathlib import Path
import pandas as pd
import plotly
import plotly.express as px
import timedelta
# import plotly.figure_factory as ff

EXCEL_FILE = Path.cwd() / "Gant_Task_Qlik.xlsx"

# Read Dataframe from Excel file
df = pd.read_excel(EXCEL_FILE)

# Assign Columns to variables
# tasks = df["Task_Name"]
# start = df["Start"]
# finish = df["End"]
#duration = df["Duration"]
################################
# fig = px.timeline(
#     df, x_start=start, x_end=finish, y=tasks, title="Task Overview"
# )
# print(df2)
#fig = px.timeline(df, x_start="Start", x_end="Finish" #, y="Resource", color="Resource"
#########################                 )
# Create Gantt Chart
# fig = px.timeline(
#     df2, x_start=start, x_end=finish, y=tasks
# )

# Upade/Change Layout
# fig.update_yaxes(autorange="reversed")
# fig.update_layout(title_font_size=42, font_size=18, title_font_family="Arial")
# plotly.offline.plot(fig, filename="Task_Overview_Gantt.html")
# #Create Gantt Chart




workbook.close()