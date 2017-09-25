import poplib
from email import parser
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr

# Input email address password and pop3 server
email = 'xxx'
password = 'xxx'
pop3_server = 'pop.gmail.com'

# Connecting to pop3 server
server = poplib.POP3_SSL(pop3_server)
# Debug messeage
# server.set_debuglevel(1)


#print(server.getwelcome())
# Idetification 
server.user(email)
server.pass_(password)
# Num of emails and space left
resp, mails, octets = server.list()

messages = [server.retr(i) for i in range(1, len(server.list()[1]) + 1)]  
messages = [b'\r\n'.join(mssg[1]).decode() for mssg in messages]  
messages = [Parser().parsestr(mssg) for mssg in messages]
for message in messages:  
        subject = message.get('Subject')      
        if subject == 'Report Delivery Notification FTP Report':
        #if from == 'alerts@greenkoncepts.com':
            for part in message.walk():  
                fileName = part.get_filename()  
                # Save the attachment  
                if fileName:  
                    with open(fileName, 'wb') as fEx:
                        data = part.get_payload(decode=True) 
                        fEx.write(data)  
                        print("Attachment %s has been saved" % fileName)
server.quit()







#index = len(mails)
#resp, lines, octets = server.retr(index)

# lines store every line of the email
# Return the original email
#msg_content = b'\r\n'.join(lines).decode('utf-8')
# Parser email
#msg = Parser().parsestr(msg_content)
# delete the emails
# server.dele(index)




# 关闭连接:            
# server.quit()

