import pandas as pd
import pyodbc

OutputDataSet1 = pd.read_excel("Output.xlsx", sheet_name = "report", header = 0, names = ["SynupID", "period time", "Date_Report", "Profile Views", "Website Views","Phone Calls","Direction Requests","Direct Searches","Discovery Searches","URL"])
OutputDataSet2 = pd.read_excel("Output.xlsx", sheet_name = "Errors", header = 0, names = ["SynupID", "Date_Report","Errors","Url"])

conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=RON\SQLEXPRESS;'
                      'Database=test_database;'
                      'Trusted_Connection=yes;')
cursor = conn.cursor()


cursor.execute('''
		CREATE TABLE products (
			SynupID nvarchar(50),
			period_time nvarchar(50),
			Date_Report nvarchar(100),
            Profile_Views nvarchar(100),
            Website_Views nvarchar(100),
            Phone_Calls nvarchar(100),
            Direction_Requests nvarchar(100),
            Direct_Searches nvarchar(100),
            Discovery_Searches nvarchar(100),
            URL nvarchar(2000)
			)
               ''')
for row in df.itertuples():
    cursor.execute('''
                INSERT INTO reports (SynupID,period_time,Date_Report,Profile_Views,Website_Views,Phone_Calls,Direction_Requests,Direct_Searches,Discovery_Searches,URL)
                VALUES (?,?,?,?,?,?,?,?,?,?)
                ''',
                   
                row.SynupID, 
                row.period_time,
                row.Date_Report,
                row.Profile_Views,
                row.Website_Views,
                row.Phone_Calls,
                row.Direction_Requests,
                row.Direct_Searches,
                row.Discovery_Searches,
                row.URL
                )
conn.commit()


cursor.execute('''
		CREATE TABLE products (
			SynupID nvarchar(50),
			Date_Report nvarchar(100),
            Errors nvarchar(200),
            Url nvarchar(2000)
			)
               ''')
for row in df.itertuples():
    cursor.execute('''
                INSERT INTO Errors (SynupID,Date_Report,Errors,Url)
                VALUES (?,?,?,?)
                ''',
                   
                row.SynupID, 
                row.Date_Report,
                row.Errors,
                row.Url
                )
conn.commit()