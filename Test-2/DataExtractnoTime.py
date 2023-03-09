import schedule
import time
import pyodbc
import csv
from datetime import datetime



# Connect to the database
conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Shiprite\Shiprite.mdb;')
cursor = conn.cursor()

# Extract the data from the table
cursor.execute("SELECT * FROM manifest")

rows = cursor.fetchall()
Cnum = conn.cursor()
Cnum.execute("SELECT PostNetCDHL_CenterID FROM Setup2 WHERE id= 1")
roz = Cnum.fetchall()


# get current date and time
current_datetime = datetime.now().strftime("%m-%d-%y")
print(current_datetime)

# convert datetime obj to string
str_current_datetime = str(current_datetime)
# Save the data to a CSV file
with open(f'C:\\Shiprite\\{roz[0][0]}_Manifest_{current_datetime}.csv', "w", newline='') as file:
    writer = csv.writer(file)
    writer.writerow([i[0] for i in cursor.description]) # write headers
    writer.writerows(rows)

# Close the connection to the database
conn.close()
print("Table exported to CSV successfully")


