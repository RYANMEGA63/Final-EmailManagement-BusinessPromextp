
import os # EmailSendAttachment
if not os.path.exists('C:\PromotexPSendMails'):
    os.mkdir('C:/PromotexPSendMails')
from pathlib import Path
from gmail_api import search_emails, init_gmail_service, get_email_messages, get_email_message_details, send_email, download_attachments_parent
import shutil
import stat

Mail_folder = r'C:\PromotexPSendMails'
Send_path = r"C:\PromotexPSendMails\EmailSend.txt"
get_path = r'C:\PromotexPSendMails\GetMails.txt'
Mails_Path = r'C:\PromotexPSendMails\Mails.txt'

def on_rm_error(func, path, exc_info):
    # Change les permissions pour pouvoir supprimer
    os.chmod(path, stat.S_IWRITE)
    func(path)

# if Mails_Path doesnt exist, create one
if not os.path.exists(Mails_Path):
    with open(Mails_Path, 'w') as file:
        file.write("0&0&")
if not os.path.exists(Send_path):
    with open(Send_path, 'w') as file:
        file.write("0&0&")   
        print('file created')     
    
if not os.path.exists(Send_path):
    with open(Send_path, 'w') as file:
        file.write('')
        
if os.path.exists(get_path):
    os.remove(get_path)

with open('C:\PromotexPSendMails\GetMails.txt', "w") as file:
    file.write("")
    
choose = 0
while choose != 'END':
    if os.path.exists(Mails_Path):
        os.remove(Mails_Path)
    if os.path.exists(Send_path):
        os.remove(Send_path)
    
    print('~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~')
    print('steps:')
    print('1. what to do: ')
    what_to_do = input('r/rf/s/sm/read ')
    
    if what_to_do == 'read':
        with open(r'C:\PromotexPSendMails\GetMails.txt', 'r') as f:
            getMail = f.read()
        
        mails = getMail.split('$')
        print(mails)
        mailN = 0
        for amail in mails:
            mailN += 1
            print(amail)
            args = amail.split('&')
            print(f'id: {args[0]}')
            print(f'subject: {args[1]}')
            print(f'sender: {args[2]}')
            print(f'reciepients: {args[3]}')
            print(f'body: {args[4]}')
            print(f'snippet: {args[5]}')
            print(f'has_attachments: {args[6]}')
            print(f'date: {args[7]}')
            print(f'star: {args[8]}')
            print(f'label: {args[9]}')
            
    
    if what_to_do == 'r' or what_to_do == 'rf':
        print('~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~')
        print('steps:')
        print('1. number of mails: ')
        NnumbMail = input('-> ')
        NFiltMail = '1'
        if what_to_do == 'rf':
            print('2. filter mail: ')
            NFiltMail = input('-> ')
            
        with open(Mails_Path, 'w') as f:
            f.write(f'{what_to_do}&{NnumbMail}&{NFiltMail}')
    
    if what_to_do == 's' or what_to_do == 'sm':
        print('~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~')
        print('steps:')
        print('1. to who: ')
        NtoWho = input('-> ')
        print('2. Subject: ')
        NSubject = input('-> ')
        print('3. Body: ')
        NBody = input('-> ')
        print('4. attachement: ')
        Nattachement = input('-> ')
        with open(Mails_Path, 'w') as f:
            f.write(f'{what_to_do}&0&0')   
        with open(Send_path, 'w') as f:
            f.write(f'{NtoWho}&{NSubject}&{NBody}&{Nattachement}')

    with open(Mails_Path, 'r') as file:
        file_componants = file.read()
        arguments = file_componants.split("&")
        
        order = arguments[0]
        num_mails = int(arguments[1])
        
        mail_filter_sender = arguments[2]
        mail_filter = f"from:{arguments[2]}"

    client_file = 'C:\PromotexPSendMails\SecretFilePromotexP.json'
    user_id = 'me'
    service = init_gmail_service(client_file)
        
    if os.path.exists('C:\PromotexPSendMails\EmailGetAttachment'):
        shutil.rmtree('C:\PromotexPSendMails\EmailGetAttachment', onerror=on_rm_error)

    os.mkdir('C:\PromotexPSendMails\EmailGetAttachment')

    def read():
        mail = 0
        messages = get_email_messages(service, max_results=num_mails) # EmailGetAttachment
        for msg in messages:
            mail += 1
            mail_id = msg['id']
            MailDir = f"C:\PromotexPSendMails\EmailGetAttachment\mailG{mail}"
            try:
                os.makedirs(MailDir)
                details = get_email_message_details(service, msg['id'])
                if details:
                    with open(r'C:\PromotexPSendMails\GetMails.txt', "a") as file:
                        file.write(f"{mail_id} & {details['subject']} & {details['sender']} & {details['reciepients']} & {details['body']} & {details['snippet']} & {details['has_attachments']} & {details['date']} & {details['star']} & {details['label']} $")
                        print('written')
                download_attachments_parent(service, user_id, mail_id, MailDir)
                print("attachement dowloaded")
                print(f'mail {mail} id {mail_id} read succefully from {details['sender']}')
            except Exception as e:
                print(f'a probleme happened when reading a mail {e}')
                
            
    def send(): 
        with open(Send_path, "r") as file:
            content = file.read() 

        arguments = content.split("&")
        
        to_address = arguments[0]
        email_subject = arguments[1]
        email_body = arguments[2]
        attachment_Files = arguments[3].split("$")
        
        mails = to_address.split("$")
        
        for i in range(len(attachment_Files)):
            if attachment_Files[i] == '':
                attachment_Files.pop(i)

        for Tomail in mails:
            print(Tomail)
            try:
                response_email_sent = send_email(
                    service,
                    Tomail,
                    email_subject,
                    email_body,
                    body_type='plain',
                    attachment_paths=attachment_Files
                )
                print(f'mail {email_subject} sent succefully to {to_address}')
            except Exception:
                print('a probleme happened when sending a mail')   

    def sendM(): #if more than one attachement, send multiple mails
        with open(Send_path, "r") as file:
            content = file.read() 

        arguments = content.split("&")
        
        to_address = arguments[0]
        email_subject = arguments[1]
        email_body = arguments[2]
        attachment_Files = arguments[3].split("$")
        
        mails = to_address.split("$")
        
        for i in range(len(attachment_Files)):
            if attachment_Files[i] == '':
                attachment_Files.pop(i)

        for Tomail in mails:
            print(Tomail)
            for attachement_single_file in attachment_Files:
                attachement_single_file = [attachement_single_file]
                
                print(attachement_single_file)
                try:
                    response_email_sent = send_email(
                        service,
                        Tomail,
                        email_subject,
                        email_body,
                        body_type='plain',
                        attachment_paths=attachement_single_file
                    )
                    print(f'mail {email_subject} sent succefully to {to_address}')
                except Exception:
                    print('a probleme happened when sending a mail') 
            
    def readF():
        mail = 0
        print('rf')
        messages = search_emails(service, mail_filter, max_results=num_mails) # EmailGetAttachment
        print('message gotten')
        for msg in messages:
            print('looking through msg')
            mail += 1
            mail_id = msg['id']
            MailDir = f"C:\PromotexPSendMails\EmailGetAttachment\mailF{mail}"
            try:
                os.makedirs(MailDir)
                print('dir made')
                details = get_email_message_details(service, msg['id'])
                print('details gotten')
                print(f"subject : {details['subject']}")
                if details:
                    with open(get_path, "a") as file:
                        file.write(f"{mail_id} &")
                        file.write(f"{details['subject'].replace('\u202f', ' ')} & {mail_filter_sender} & {details['body'].replace('\u202f', ' ')} & {details['snippet'].replace('\u202f', ' ')} & {details['date'].replace('\u202f', ' ')} & {details['label'].replace('\u202f', ' ')} $")
                download_attachments_parent(service, user_id, mail_id, MailDir)
                print(f'mail {mail} id {mail_id} read succefully from {details['sender']}')
            except Exception:
                print('a probleme happened when sending a mail')  
                
    if order == "r":     
        read()       
    elif order == "s":
        send()     
    elif order == "sm":
        sendM()     
    elif order == "rf":
        readF()        

    else:
        print('not a valid operation') 
