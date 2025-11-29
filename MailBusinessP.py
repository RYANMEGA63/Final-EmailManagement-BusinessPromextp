try:
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
            file.write('')
            
    if os.path.exists(get_path):
        os.remove(get_path)

    with open('C:\PromotexPSendMails\GetMails.txt', "w") as file:
        file.write("")

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
            except Exception:
                with open(Mails_Path, 'a') as file:
                    print(f"error in {MailDir} ")
                    file.write('x&')
                
        with open(Mails_Path, 'a') as file:
            file.write('1&')
            
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
            except Exception:
                with open(Mails_Path, 'a') as file:
                    file.write('x&')
        with open(Mails_Path, 'a') as file:
            file.write('1&')     

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
                except Exception:
                    with open(Mails_Path, 'a') as file:
                        file.write('x&')
        with open(Mails_Path, 'a') as file:
            file.write('1&')
            
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
            except Exception:
                with open(Mails_Path, 'a') as file:
                    file.write('x&')
                
        with open(Mails_Path, 'a') as file:
            file.write('1&')
    if order == "r":     
        read()       
    elif order == "s":
        send()     
    elif order == "sm":
        sendM()     
    elif order == "rf":
        readF()        

    else:
        with open(Mails_Path, 'a') as file:
            file.write('0&')
except Exception:
    with open(Mails_Path, "a") as file:
        file.write("0&")
            
        