try:
    import os # EmailSendAttachment
    from pathlib import Path
    from gmail_api import search_emails, init_gmail_service, get_email_messages, get_email_message_details, send_email, download_attachments_parent
    import shutil
    import stat
    
    def on_rm_error(func, path, exc_info):
        # Change les permissions pour pouvoir supprimer
        os.chmod(path, stat.S_IWRITE)
        func(path)
    
    if not os.path.exists('C:\Mails.txt'):
        with open('C:\Mails.txt', 'w') as file:
            file.write("0&0&")
            
    if not os.path.exists('C:\EmailSendAttachment'):
        os.mkdir('C:/EmailSendAttachment')
        
    if not os.path.exists("C:\EmailSend.txt"):
        with open("C:\EmailSend.txt", 'w') as file:
            file.write('')
    
    with open('C:\Mails.txt', 'r') as file:
        file_componants = file.read()
        arguments = file_componants.split("&")
        
        order = arguments[0]
        num_mails = int(arguments[1])
        
        mail_filter_sender = arguments[2]
        mail_filter = f"from:{arguments[2]}"
    
    client_file = 'C:\SecretFilePromotexP.json'
    user_id = 'me'
    service = init_gmail_service(client_file)
    
    if os.path.exists('C:\GetMails.txt'):
        os.remove('C:\GetMails.txt')
        
    if os.path.exists('C:\EmailGetAttachment'):
        shutil.rmtree('C:\EmailGetAttachment', onerror=on_rm_error)

    os.mkdir('C:\EmailGetAttachment')
    
    mail = 0
    
    if order == "r":     
        messages = get_email_messages(service, max_results=num_mails) # EmailGetAttachment
        for msg in messages:
            mail += 1
            mail_id = msg['id']
            MailDir = f"C:\EmailGetAttachment\mailG{mail}"
            try:
                os.makedirs(MailDir)
                details = get_email_message_details(service, msg['id'])
                if details:
                    details = details.replace('\u202f', ' ')
                    with open('C:\GetMails.txt', "a") as file:
                        file.write(f"{mail_id} &")
                        f"{details['subject']} & {details['sender']} & {details['reciepients']} & {details['body']} & {details['snippet']} & {details['has_attachement']} & {details['date']} & {details['star']} & {details['label']} $"
                download_attachments_parent(service, user_id, mail_id, MailDir)
            except Exception:
                with open('C:\Mails.txt', 'a') as file:
                    file.write('x&')
               
        with open('C:\Mails.txt', 'a') as file:
            file.write('1&')
            
    elif order == "s":
        file_path = r"C:\EmailSend.txt"
        with open(file_path, "r") as file:
            content = file.read() 
         
        arguments = content.split("&")
        
        to_address = arguments[0]
        email_subject = arguments[1]
        email_body = arguments[2]
        
        attachment_dir = Path('C:\EmailSendAttachment')
        attachment_Files = list(attachment_dir.glob('*'))

        response_email_sent = send_email(
            service,
            to_address,
            email_subject,
            email_body,
            body_type='plain',
            attachment_paths=attachment_Files
        )
        with open('C:\Mails.txt', 'a') as file:
            file.write('1&')
    
    elif order == "rf":
        messages = search_emails(service, mail_filter, max_results=num_mails) # EmailGetAttachment
        for msg in messages:
            mail += 1
            mail_id = msg['id']
            MailDir = f"C:\EmailGetAttachment\mailF{mail}"
            try:
                os.makedirs(MailDir)
                details = get_email_message_details(service, msg['id'])
                print(f"subject : {details['subject']}")
                if details:
                    with open('C:\GetMails.txt', "a") as file:
                        file.write(f"{mail_id} &")
                        file.write(f"{details['subject'].replace('\u202f', ' ')} & {mail_filter_sender} & {details['body'].replace('\u202f', ' ')} & {details['snippet'].replace('\u202f', ' ')} & {details['date'].replace('\u202f', ' ')} & {details['label'].replace('\u202f', ' ')} $")
                download_attachments_parent(service, user_id, mail_id, MailDir)
            except Exception:
                with open('C:\Mails.txt', 'a') as file:
                    file.write('x&')
               
        with open('C:\Mails.txt', 'a') as file:
            file.write('1&')        
    
    else:
        with open('C:\Mails.txt', 'a') as file:
            file.write('0&')
except Exception:
    with open('C:\Mails.txt', "a") as file:
        file.write("0&")
        
        