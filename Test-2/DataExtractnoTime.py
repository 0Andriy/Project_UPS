import argparse
import csv
import os
import pysftp
import pyodbc
import schedule
import time
from datetime import datetime, timedelta
from getpass import getpass

# Set up command line arguments
parser = argparse.ArgumentParser(description='Export data from an Access database to a CSV file and upload it to an FTP server.')
parser.add_argument('--all', '-all', action='store_true', help='export data for all days instead of just today')
parser.add_argument('--savefiles', '-savefiles', action='store_true', help='do not delete the exported file after uploading it to the FTP server')
args = parser.parse_args()

# Connect to the database
conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\PC\Desktop\Project_UPS\Test-2\TX251 Shiprite.mdb;')
cursor = conn.cursor()

# Extract the data from the table
if args.all:
    # Export all data
    cursor.execute("SELECT * FROM manifest")
else:
    # Export data for today
    today = datetime.now().date()
    # Test
    # today = datetime(2022,6,13) 
    # print(today)
    tomorrow = today + timedelta(days=1)
    # print(tomorrow)
    cursor.execute(f"SELECT * FROM manifest WHERE Date >= #{today}# AND Date < #{tomorrow}#")

rows = cursor.fetchall()
Cnum = conn.cursor()
Cnum.execute("SELECT PostNetCDHL_CenterID FROM Setup2 WHERE id= 1")
roz = Cnum.fetchall()[0][0]

# Get current date and time
current_datetime = datetime.now().strftime("%m-%d-%y")

# Save the data to a CSV file
filename = f'{roz}_Manifest_{current_datetime}.csv'
with open(os.path.join(filename), "w", newline='', encoding='utf8') as file:
    writer = csv.writer(file)
    writer.writerow([i[0] for i in cursor.description]) # Write headers
    writer.writerows(rows)

# Close the connection to the database
conn.close()
print("Table exported to CSV successfully")

# Connect to the FTP server
ftp_server = 'ftp.dash-ipostnet.com'
ftp_username = 'manifest_file_upload@dash-ipostnet.com'
ftp_password = None

while ftp_password is None:
    try:
        ftp_password = getpass(prompt='Enter the FTP password: ')
        with pysftp.Connection(host=ftp_server, username=ftp_username, password=ftp_password, port=21) as ftp:
            # Change to the appropriate directory
            ftp.cwd('/manifest_files')

            # Upload the file
            ftp.put(os.path.join('C:\\Shiprite', filename))
            print(f"File {filename} uploaded to FTP server successfully")
    except pysftp.AuthenticationException:
        print("Incorrect FTP password. Please try again.")
        ftp_password = None

# Delete the file if the user did not request to save it
if not args.savefiles:
    os.remove(os.path.join(filename))
    print(f"File {filename} deleted from local directory")
