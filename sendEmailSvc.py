import smtplib, os, html2text, win32serviceutil , debugpy, requests, uuid, configparser
from email.utils import formataddr
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from pathlib import Path
from MedatechUK.APY.oDataConfig import Config
from MedatechUK.APY.mLog import mLog
from MedatechUK.APY.cl import folderWatch , clArg
from MedatechUK.APY.svc import AppSvc
from server import Server

class MySVC(AppSvc):    
    _svc_name_ = "sendEmailSVC"
    _svc_display_name_ = "Send Emails"    

    def __init__(self , args):    
        self.Main = MySVC.main   
        self.Init = MySVC.init   
        self.Folder = Path(__file__).parent         
        AppSvc.__init__(self , args)

    def init(self):
        if self.debuginit: debugpy.breakpoint()         
        pass

    def emailUpdateStatement(self,server,company,id,sent):
        return  (
            "USE [{}];".format(company)
            + " UPDATE {}.{}.dbo.ZGEM_EMAILING".format(server,company)
            + " SET SENT = {}".format(sent)
            + " , SENT_DATE = DATEDIFF(MI,'01/01/1988',GETDATE())"
            + " WHERE EMAILID = {}".format(id)
        )

    def getEmailTemplateStatement(self, server,company, exec, num, emailid, use_manual_emailbody):
        email_template_query = """
                SELECT EXTMSGTEXT.TEXT
                FROM {}.system.dbo.EXTMSGTEXT EXTMSGTEXT
                WHERE EXTMSGTEXT.T$EXEC = {}
                AND   EXTMSGTEXT.NUM = {}
                AND EXTMSGTEXT.TEXTLINE <> 0
                AND EXTMSGTEXT.T$EXEC <> 0
                AND   EXTMSGTEXT.NUM <> 0
                ORDER BY TEXTORD
                """.format(server,exec,num)
        if use_manual_emailbody:
            email_template_query = """
                SELECT ZGEM_EMAILBODY.TEXT
                FROM {}.{}.dbo.ZGEM_EMAILBODY ZGEM_EMAILBODY
                WHERE ZGEM_EMAILBODY.EMAILID = {}
                AND ZGEM_EMAILBODY.TEXTLINE <> 0
                AND ZGEM_EMAILBODY.EMAILID <> 0
                ORDER BY TEXTORD
                """.format(server,company,emailid)

        return email_template_query

    def getAttachmentStatement(self,server,company,id): 
        return (
            "SELECT EXTFILENAME, NAME"
            + " FROM {}.{}.dbo.ZGEM_EMAILATTCHMNTS".format(server,company)
            + " WHERE EMAILID = {}".format(id)
            + " UNION"
            + " SELECT EXTFILENAME, ATTACHMENTNAME"
            + " FROM {}.{}.dbo.ZGEM_EMAILING E,".format(server,company)
            + " {}.system.dbo.ZGEM_DEFATTACHMNTS A".format(server)
            + " WHERE E.EMAILID = {}".format(id)
            + " AND E.TEMPLATE_EXEC = A.T$EXEC"
            + " AND E.TEMPLATE_NUM = A.NUM;"
        )

    def getVariablesStatement(self,server,company,id):
        return (
            "SELECT VARIABLENAME, VALUE"
            + " FROM {}.{}.dbo.ZGEM_EMAILVARIABLES".format(server,company)
            + " WHERE EMAILID = {}".format(id)
        )
    
    def getNextEmailInEmailStack(self, server, company):
        return """
                SELECT TOP 1 *
                FROM {}.{}.dbo.ZGEM_EMAILING
                WHERE EMAILID <> 0
                AND READY_TO_SEND = 'Y'
                AND SENT <> 'Y'
                AND RECIPIENT_EMAIL <> ''
                AND SENT_DATE <= DATEDIFF(MI,'01/01/1988',GETDATE() - 1)
                AND DATEDIFF(MI,'01/01/1988',GETDATE()) - CURDATE >= 1
                ORDER BY CURDATE ASC
                """.format(server, company)

    def main(self):       
        if self.debug: debugpy.breakpoint # -debug          
        #Get Working Directory
        WorkingDir = Path(__file__).parent
        WorkingDir = str(WorkingDir).replace('\\', "\\\\")

        #Default Template
        defaultHtmlEmailBody = """
        <html>
        <head></head>
        <body>
            <p>Hi!<br>
            There was no email template defined.
            </p>
        </body>
        </html>
        """
        #IT Send As Permission Setup Request
        IT_send_as_setup_html = """
        <html>
        <head></head>
        <body>
            <p>Hi!<br>
            Please set up send as permisisng for the following.<br>
            Allow <SENDFROM> to send as <SENDAS>.
            </p>
        </body>
        </html>
        """
        #STMP Email Settings
        settingsINI = configparser.ConfigParser()
        settingsINI.read(str(Path(__file__).parent) + '\\settings.ini')
        IT_email = settingsINI['SMTPSettings']['ITemail']
        Priority_email = settingsINI['SMTPSettings']['email']
        Priority_password = settingsINI['SMTPSettings']['password']
        smtp_server = settingsINI['SMTPSettings']['server']
        smtp_port = settingsINI['SMTPSettings']['port']

        live = Server('Live','[CE-AZ-UK-S-PRIO\\PRI]', ['base'], 'https://priority.clarksonevans.co.uk/', settingsINI['LIVE']['server'],settingsINI['LIVE']['credentials'])
        dev = Server('Dev','[CE-AZ-UK-S-PRIO\\DEV]', ['test'], 'https://prioritydev.clarksonevans.co.uk/', settingsINI['DEV']['server'],settingsINI['DEV']['credentials'])
        test = Server('Test','[CE-AZ-UK-S-PRIO\\TST]', ['base', 'test'], 'https://prioritytest.clarksonevans.co.uk/', settingsINI['TEST']['server'],settingsINI['TEST']['credentials'])
        servers = [live, test, dev]
        for server in servers:
            for company in server.companies:
                #Connect to Database
                config = Config(
                    env = company
                    , path = WorkingDir
                )
                nextEmailInStack = self.getNextEmailInEmailStack(server.server, company)
                cnxn = Config.zgem_cnxn(config, server.serverIP, server.serverCredentials)
                cnxn2 = Config.zgem_cnxn(config, server.serverIP, server.serverCredentials)
                while True:
                    cursor1 = cnxn.cursor()
                    cursor1.execute("USE [{}]; ".format( company ))
                    emails = cursor1.execute(nextEmailInStack)
                    if emails.rowcount == 0:
                        break
                    for email in emails:
                        #Get email variables
                        id = email[0]
                        sender_name = email[2]
                        sender_email = email[3]
                        recipient_email = email[4]
                        subject = email[5]
                        exec = email[6]
                        num = email[7]
                        emailType = email[10]
                        use_manual_emailbody = True if email[11].upper() == 'Y' else False
                        #Get Email Template
                        getEmailTemplate = self.getEmailTemplateStatement(server.server, company, exec, num, id, use_manual_emailbody)
                        emailTemplate = cursor1.execute(getEmailTemplate)
                        html = defaultHtmlEmailBody
                        if emailTemplate.rowcount != 0:
                            html = ''
                            prevLine = ''
                            for line in emailTemplate:
                                gap = ' '
                                if 'src' in prevLine or 'href' in prevLine or 'https' in prevLine:
                                    if prevLine.count('"') % 2 != 0:
                                        gap = ''
                                html += gap + line[0]
                                prevLine = line[0]
                        #Replace variables in email template
                        getVariables = self.getVariablesStatement(server.server,company,id)
                        variables = cursor1.execute(getVariables)
                        if variables.rowcount != 0:
                            for variable in variables:
                                html = html.replace(variable[0], variable[1])
                        #Build Plain&Html Email Body
                        msg_alternative = MIMEMultipart('alternative')
                        text_part = MIMEText(html2text.html2text(html), 'plain')
                        msg_alternative.attach(text_part)
                        self.log.logger.debug ("HTML message body: " + html)
                        html_part = MIMEText(html, 'html')
                        msg_alternative.attach(html_part)
                        #Define Email variables
                        message = MIMEMultipart('mixed')
                        message['Subject'] = subject
                        message['From'] = formataddr((sender_name, sender_email))
                        message['To'] = recipient_email
                        message.attach(msg_alternative)
                        #Get attachments
                        getAttachments = self.getAttachmentStatement(server.server, company,id)
                        attachments = cursor1.execute(getAttachments)
                        if attachments.rowcount != 0:
                            for attachment in attachments:
                                if attachment[0]:
                                    fileName = attachment[1]
                                    url = server.domain + attachment[0].replace("../","").replace('system/','').replace('mail/','primail/')
                                    self.log.logger.debug("Attachment: " + attachment[0])
                                    self.log.logger.debug("Attachment URL: " + url)
                                    path, fileExtension = os.path.splitext(url)
                                    response = requests.get(url)
                                    newFile = str(Path(__file__).parent) + '/'+str(uuid.uuid4())+fileExtension
                                    with open(newFile, 'wb') as file:
                                        file.write(response.content)
                                    with open(newFile,'rb') as file:
                                        # Attach the file with filename to the email
                                        message.attach(MIMEApplication(file.read(), Name = fileName + fileExtension))
                                    os.remove(newFile)
                        #Send Email
                        emailWasSent = 0
                        try:
                            with smtplib.SMTP(smtp_server, smtp_port) as SMTPserver:
                                SMTPserver.starttls()
                                SMTPserver.login(Priority_email, Priority_password)
                                SMTPserver.sendmail(Priority_email, recipient_email, message.as_string())
                                emailWasSent = 1
                                self.log.logger.debug("Server: {} ID: {} Type: {} From: {} To: {} Subject: {}".format(server.name,id,emailType,sender_email,recipient_email,subject))
                        except Exception as ex:
                            if type(ex).__name__ == "SMTPDataError":
                                if "SendAsDenied" in str(ex.args[1]):
                                    try:
                                        with smtplib.SMTP(smtp_server, smtp_port) as SMTPserver1:
                                            #Reset sender and recipient
                                            del message["From"]
                                            del message["To"]
                                            message["From"] = formataddr(('Clarkson Evans ERP System', Priority_email))
                                            message["To"] = recipient_email + ", " + sender_email
                                            message.add_header('reply-to', sender_email)
                                            #Send email from Priority email to recipient and sender
                                            SMTPserver1.starttls()
                                            SMTPserver1.login(Priority_email, Priority_password)
                                            SMTPserver1.sendmail(Priority_email, [recipient_email, sender_email], message.as_string())
                                            emailWasSent = 1
                                            self.log.logger.debug("Server: {} ID: {} Type: {} ({} can't send as {} but email was still sent) From: {} To: {},{} Subject: {}".format(server.name,id,emailType,Priority_email,sender_email,Priority_email,recipient_email,sender_email,subject))
                                            #Notify IT to set up send as permission
                                            sendAsMessage = MIMEMultipart('alternative')
                                            IT_html = IT_send_as_setup_html.replace("<SENDFROM>",Priority_email).replace("<SENDAS>",sender_email)
                                            text_part = MIMEText(html2text.html2text(IT_html), 'plain')
                                            sendAsMessage.attach(text_part)
                                            html_part = MIMEText(IT_html, 'html')
                                            sendAsMessage.attach(html_part)
                                            sendAsMessage['Subject'] = "Set Up Send As Permission"
                                            sendAsMessage['From'] = formataddr(('Clarkson Evans ERP System', Priority_email))
                                            sendAsMessage['To'] = IT_email
                                            SMTPserver1.sendmail(Priority_email, IT_email, sendAsMessage.as_string())
                                    except Exception as ex1:
                                        errorMsg = "Server: {0} SMTP Email Connection Failed Error Type: {1}, Arguments: {2!r}".format(server.name,type(ex1).__name__, ex1.args)
                                        self.log.logger.critical(errorMsg)
                            else:
                                errorMsg = "Server: {0} SMTP Email Connection Failed  Error Type: {1}, Arguments: {2!r}".format(server.name,type(ex).__name__, ex.args)
                                self.log.logger.critical(errorMsg)
                    #Update Priority Database
                    sent = "'Y'" if emailWasSent == 1 else "'N'"
                    update = self.emailUpdateStatement(server.server,company,id,sent)
                    cursor2 = cnxn2.cursor()
                    cursor2.execute(update)
                    cursor2.commit()
                    cursor2.close()
                    cursor1.close() 
        pass

if __name__ == '__main__':    
    win32serviceutil.HandleCommandLine(MySVC)
