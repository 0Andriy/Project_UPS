from ftplib import FTP_TLS

def ftp_exp(ftp_usr, ftp_pass, port):
    ftp = FTP_TLS()
    ftp.set_debuglevel(2)
    ftp.connect(host='ftp.EX.com', port=port)
    ftp.login(user=ftp_usr, passwd=ftp_pass)
    print(ftp.getwelcome())
    ftp.storbinary('STOR image.jpg', open('image.jpg','rb'))
    print(ftp.dir())
    ftp.close()

if __name__ == "__main__":
    ftp_exp(port=21, ftp_usr='manifest_file_upload@dash-ipostnet.com', ftp_pass='SC{{m@$PZ2;L')
