import smtplib, ssl
import time
import re
import datetime
import csv
from string import Template
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import sys
from pathlib import Path


#-----
myEmail = "enterEmail"
myPass = "enterPassword"
myName = "enterName " + "<" + myEmail + ">"
CompanySMTP = "enterCompanySMTP"
#-----


#PREBERI CSV, ljudi, ki jim bomo poslali
def get_contacts(filename):
	try:
		with open(filename) as fp:
			reader = csv.reader(fp, delimiter=";")
			next(reader, None)  # skip the headers
			return [row for row in reader]
	except FileNotFoundError:
		print("Error: Manjka datoteka emails.csv.")
		input("Press enter to exit....")
		sys.exit(0)
	except:
		print("Error: Prišlo je do napake pri branju datoteke emails.csv.")
		input("Press enter to exit....")
		sys.exit(0)


#PREBERI SPOROCILO, ki jo bomo poslali
def read_template(filename):
	try:
		with open(filename, 'r', encoding='utf-8') as template_file:
			template_file_content = template_file.read()
		return Template(template_file_content)
	except FileNotFoundError:
		print("Error: Manjka datoteka vsebina.txt.")
		input("Press enter to exit....")
		sys.exit(0)
	except:
		print("Error: Prišlo je do napake pri branju vsebine maila.")
		input("Press enter to exit....")
		sys.exit(0)


#Ask for input and check if it is number
def get_number():
	gotNo = False
	while not gotNo:
		#ask for input
		number = input()
		#check if input is number
		try:
			number = int(number)
			gotNo = True
		except:
			print("Vneseno ni številka. Prosim poskusite še enkrat...")
	return number


def SMTP_connect():
    #Connect to SMTP
    global s
    s = smtplib.SMTP(CompanySMTP, 26)
    s.starttls()
    #try connecting to server
    try:
        print("Connecting...")
        s.login(myEmail, myPass)
    except smtplib.SMTPAuthenticationError:
        #geslo je napacno
        print("Error: Napačno geslo")
        input("Press enter to exit....")
        sys.exit(0)
    except:
        #se nemore povezat
        print("Error: Connection problems.")
        input("Press enter to exit....")
        sys.exit(0)	
    else:
        #povezava uspesna
        print("Connection successful")


def personalised_hi(contact_fName,contact_lName,contact_company):
    if contact_fName!="" and contact_lName!="":
        #we know both first name and last name
        hi = "Dear " + contact_fName + " " + contact_lName + ","
    elif contact_lName!="":
        #we know only last name
        #hi = "Dear " + contact_lName + ","
        hi = "Dear Sir/Madam,"
    else:
        if contact_company!="":
        	#we know only the company name
        	hi = "Hi, " + contact_company +" team!"
        else:
        	#we don't know anything
        	hi = "Hi!"
    return hi


#read files and set variables
data_read = get_contacts("emails.csv")
message_template = read_template('vsebina.txt')
print("Got the message and contacts")
l = len(data_read)	#length of rows in csv
i = 0				#sesteva emaile za izpis
hi = ""				#za intro sporocilo
sent_data = []		#komu smo že vse poslali
imeDat = ""			#attachemnt, če obstaja
 
#Ask for subject:
print("Kaj naj bo zadeva?")
subject = input()


#Ask for package number
print("Po kolko mailov jih naj poslje naenkrat?")
packageEmail = get_number()
#package can not be 0 - check if 0
while True:
	if packageEmail==0:
		print("Paket nemore vsebovat 0 mailov. Prosim poskusite še enkrat...")
		packageEmail = get_number()
		continue
	break


#Ask for Time sleep number
print("Kakšna je naj pavza med njimi (v sekundah)?")
timeSleep = get_number()


#Ask for attachment:
gotAttacthment = False
while not gotAttacthment:
	#ask for file
	imeDat = input("Točno ime attachementa (če ga ni preskoči ta del):\n")
	#if empty - no attachement
	if imeDat=="":
		break
	#check if file exists
	try:
		my_file = Path(imeDat)
		if my_file.is_file():
			gotAttacthment = True
			print("Attachement got!\n")
		else:
			print("Datoteka ne obstaja. Poskusite ponovno.")
	except:
		print("Error: Prišlo je do napake pri branju datoteke.")
		input("Press enter to exit....")
		sys.exit(0)


SMTP_connect()	#connets to mail server
# For each contact, send the email:
for contact in data_read:
	#read contact
    contact_fName = contact[0]		#read first name
    contact_lName = contact[1]		#read last name
    contact_company = contact[2]	#read company name
    contact_email = contact[3]		#read email


    #check if email is valid. If not skip it 
    if not (re.match(r"^[A-Za-z0-9\.\+_-]+@[A-Za-z0-9\._-]+\.[a-zA-Z]*$", contact_email)):
    	print("Skipping: ",contact_email, " - Invalid e-mail.")
    	continue

    hi = personalised_hi(contact_fName,contact_lName,contact_company)
    msg = MIMEMultipart()       # create a message
    message = message_template.substitute(INTRO=hi)

    # setup the parameters of the message
    msg['From']= myName
    msg['To']= contact_email
    msg['Subject']= subject

    # add in the message body
    #msg.attach(MIMEText(message, 'plain'))	#for a normal text
    msg.attach(MIMEText(message, 'html'))	#for a html text
    
    #dodaj zraven attachment ce obstaja:
    if gotAttacthment==True:
        part = MIMEBase('application', "octet-stream")
        part.set_payload(open(imeDat, "rb").read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment', filename=imeDat)
        msg.attach(part)
    
    # send the message via the server set up earlier. At least try it
    try:  
        s.send_message(msg)
        #dodaj sporočilo med "poslane" in dodaj datum kdaj se je to zgodilo
        contact.append(datetime.datetime.now())  
        sent_data.append(contact)

        #message sent, let the user know
        print("message sent to ",contact_email)  
        i = i + 1
    except:
        print("Error: Failed to send an email to ",contact_email)
        del msg
        continue     
    
    del msg

    #Ce je stevilo v "paketu" dosezeno, pocak x stevilo sekund
    if i%packageEmail==0:
        print("Progress: " , i , "/" , l)
        if i!=l:
            s.quit()	#has to brak connection 
            time.sleep(timeSleep)	#wait
            print("Sleeping...")
            SMTP_connect()	#ponovno zazeni internetno povezavo

s.quit()
print("\nFinished!\nSent:" , i , "/" , l)


#shrani seznam vseh ko emailov ko jih je ratalo poslat
with open("sentEmails.csv", "wt", newline="") as fp:
    writer = csv.writer(fp, delimiter=";")
    writer.writerow(["first name", "last name", "company", "email", "date sent"])  # write header
    writer.writerows(sent_data)

input("Press enter to exit....")