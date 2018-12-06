import smtplib 
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase 
from email import encoders 
import pandas as pd
import qrcode
from unidecode import unidecode
from tkinter import Tk
from tkinter.filedialog import askopenfilename

defaults= open("default.txt","r")
defaults=defaults.readlines()
default_mail=defaults[1][:-1]
default_password=defaults[3][:-1]
default_from=defaults[5][:-1]
def_topic=defaults[7][:-1]
class ask:
    def __init__(self,text,default_answer):
        self.x=input(text+"")
        if not self.x:
            self.x= default_answer
#you can create a varible easily with variable=ask(question,default_answer)
# in order to ask user change default settings. And you can use the variable with variable.x


print("CHOOSE THE csv FILE WHICH INCLUDE THOSE WHO YOU WANT TO SEND MAIL")
Tk().withdraw()
inputfile = askopenfilename()
y = pd.read_csv(inputfile)
print(y.head())
defaults_on_excel=['NAME', 'SURNAME',"TEL","MAIL"]
for i in range(4):
    print(" *",defaults_on_excel[i],"* Write the equivalent with it: If it is same, leave empty and enter")
    temp=input()
    if temp:    
        defaults_on_excel[i] = temp
    temp=None
print(defaults_on_excel)
y = y.filter(items=defaults_on_excel)
print(y.head())
ycsv = y.to_csv(header=False, index=False).split('\n')
print(ycsv)

fromaddr=ask("What account do you want to send mail from ," +default_mail+"'dan göndermek için boş bırakınız\n",default_mail)

password=ask("Write your password or enter to use default password\n",default_password)

# creates SMTP session 
s = smtplib.SMTP('smtp.gmail.com', 587) 
print("please wait signing...")
s.starttls() 
  
s.login(fromaddr.x, password.x) 

kimden=ask("Write FROM that receiver see\n",default_from)

konu=ask("What is TOPIC\n",def_topic)
# badi=ask("(Advanced )     Body kısmını değiştirebilirsiniz, varsayılan için kodları inceleyin, entera basıp varsayılanı kullanabilirsiniz\n ",""""Merhaba """ + name + """ \nYorum Message""")
  
errors=[]
error_count=0
QrOrNot=ask("write 0 to send without QR",1)

total_people=len(ycsv)
for i in range(len(ycsv)-1):
    
    msg = MIMEMultipart() 
    msg['From'] = kimden.x

    
# storing the subject  
    msg['Subject'] = konu.x
    name = ycsv[i].split(',')[0]
    surname = ycsv[i].split(',')[1]

    kime = ycsv[i].split(',')[3]

    # string to store the body of the mail 
    body= """Hello """ + name + """ \nYour Message \n Your Name"""
    #  # body = badi.x TO DO - ask users to change body at console


# attach the body with the msg instance 
    msg.attach(MIMEText(body, 'plain')) 

    if QrOrNot.x==1:
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=10,
            border=4,
        )

        qr.add_data(ycsv[i])
        qr.make(fit=True)
        print("\nqr created")
        img = qr.make_image(fill_color="black", back_color="white")
        
    
        img.save('./' + name + surname+konu.x+".png") 
        # 
        filename = name+surname+konu.x + ".png"

        
# open the file to be sent  

        attachment = open(filename, "rb") 
        print(filename)
        p = MIMEBase('application', 'octet-stream') 
        #try:
        p.set_payload((attachment).read()) 

        encoders.encode_base64(p) 
    
        p.add_header('Content-Disposition', "attachment", filename=('utf-8', '', filename))
        #except TypeError:
            # print("-----HATA-----")
            # errors.append(ycsv[i])
            # error_count+=1
        msg.attach(p) 

   
    
    msg['To'] = kime 
    # storing the receivers email address  
    
    #s.sendmail(fromaddr, i, text) ---- it seems to receivers as BCC , below method is more useful and advanced
    try:    
        s.send_message(msg, fromaddr.x, kime)
    except (TypeError,smtplib.SMTPRecipientsRefused):
        print("-----ERROR-----")
        errors.append(ycsv[i])
        error_count+=1
    #smtplib.SMTPRecipientsRefused or    ,,  bu hata hata verdiriyor
    print(kime, "", "    Remain mails: ",(total_people-i-2),"","   ERRORS:",error_count)
print("\n those ERRORs occured: ",errors)
print("amount of sent message:", total_people-1-error_count)
s.quit() 
