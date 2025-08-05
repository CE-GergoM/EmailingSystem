import smtplib, os, html2text, win32serviceutil , debugpy, requests, uuid, configparser, logging
from email.utils import formataddr
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from pathlib import Path
from MedatechUK.APY.oDataConfig import Config
from MedatechUK.APY.svc import AppSvc
from server import Server
from datetime import datetime

class MySVC(AppSvc):    
    _svc_name_ = "sendEmailSVC"
    _svc_display_name_ = "Send Emails"    

    def __init__(self , args, working_dir=None, config_path=None):    
        self.Main = MySVC.main   
        self.Init = MySVC.init   
        self.Folder = Path(__file__).parent
        # Set working directory
        self.working_dir = working_dir or str(Path(__file__).parent).replace('\\', "\\\\")
        
        # Set up logging
        self.setup_logging()
        
        # Load settings
        self.load_settings(config_path)
        
        # Initialize servers
        self.setup_servers()
        
        # Email templates
        self.setup_email_templates()       
        AppSvc.__init__(self , args)

    def init(self):
        if self.debuginit: debugpy.breakpoint()         
        pass

    def setup_logging(self):
        """Set up logging into log/YYYY-MM/YYMMDD.log format"""

        # Create dated log directory and file
        today = datetime.today()
        subdir = today.strftime("%Y-%m")         # "2025-08"
        filename = today.strftime("%y%m%d.log")  # "250805.log"
        
        log_dir = os.path.join(self.working_dir, "log", subdir)
        os.makedirs(log_dir, exist_ok=True)

        log_file = os.path.join(log_dir, filename)

        # File handler (logs to today's file)
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        file_handler.setLevel(logging.INFO)

        # Console handler (same format)
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        console_handler.setLevel(logging.INFO)

        # Set up logger
        self.logger = logging.getLogger('sendEmailSVC')
        self.logger.setLevel(logging.INFO)
        self.logger.propagate = False  #prevent double logging to root logger
        self.logger.handlers.clear()
        self.logger.addHandler(file_handler)
        self.logger.addHandler(console_handler)
    
    def load_settings(self, config_path=None):
        """Load settings from INI file"""
        settings_path = config_path or str(Path(__file__).parent) + '\\settings.ini'
        self.settings = configparser.ConfigParser()
        self.settings.read(settings_path)
        
        # SMTP Settings
        self.IT_email = self.settings['SMTPSettings']['ITemail']
        self.priority_email = self.settings['SMTPSettings']['email']
        self.priority_password = self.settings['SMTPSettings']['password']
        self.smtp_server = self.settings['SMTPSettings']['server']
        self.smtp_port = self.settings['SMTPSettings']['port']
    
    def setup_servers(self):
        """Initialize server configurations"""
        self.live = Server(
            'Live', 
            '[CE-AZ-UK-S-PRIO\\PRI]', 
            'https://priority.clarksonevans.co.uk/', 
            self.settings['LIVE']['server'],
            self.settings['LIVE']['credentials']
        )
        self.dev = Server(
            'Dev', 
            '[CE-AZ-UK-S-PRIO\\DEV]', 
            'https://prioritydev.clarksonevans.co.uk/', 
            self.settings['DEV']['server'],
            self.settings['DEV']['credentials']
        )
        self.test = Server(
            'Test', 
            '[CE-AZ-UK-S-PRIO\\TST]', 
            'https://prioritytest.clarksonevans.co.uk/', 
            self.settings['TEST']['server'],
            self.settings['TEST']['credentials']
        )
        
        # Default to dev server for testing
        self.servers = [self.dev, self.test, self.live]
    
    def setup_email_templates(self):
        """Set up default email templates"""
        self.default_html_email_body = """
        <html>
        <head></head>
        <body>
            <p>Hi!<br>
            There was no email template defined.
            </p>
        </body>
        </html>
        """
        
        self.IT_send_as_setup_html = """
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
    
    def set_servers(self, servers):
        """Set which servers to process"""
        self.servers = servers
    
    def email_update_statement(self, server, company, id, sent):
        return (
            "USE [{}];".format(company)
            + " UPDATE {}.{}.dbo.ZGEM_EMAILING".format(server, company)
            + " SET SENT = {}".format(sent)
            + " , SENT_DATE = DATEDIFF(MI,'01/01/1988',GETDATE())"
            + " WHERE EMAILID = {}".format(id)
        )

    def get_email_template_statement(self, server, company, exec, num, emailid, use_manual_emailbody):
        email_template_query = """
                SELECT EXTMSGTEXT.TEXT
                FROM {}.system.dbo.EXTMSGTEXT EXTMSGTEXT
                WHERE EXTMSGTEXT.T$EXEC = {}
                AND   EXTMSGTEXT.NUM = {}
                AND EXTMSGTEXT.TEXTLINE <> 0
                AND EXTMSGTEXT.T$EXEC <> 0
                AND   EXTMSGTEXT.NUM <> 0
                ORDER BY TEXTORD
                """.format(server, exec, num)
        if use_manual_emailbody:
            email_template_query = """
                SELECT ZGEM_EMAILBODY.TEXT
                FROM {}.{}.dbo.ZGEM_EMAILBODY ZGEM_EMAILBODY
                WHERE ZGEM_EMAILBODY.EMAILID = {}
                AND ZGEM_EMAILBODY.TEXTLINE <> 0
                AND ZGEM_EMAILBODY.EMAILID <> 0
                ORDER BY TEXTORD
                """.format(server, company, emailid)
        return email_template_query

    def get_attachment_statement(self, server, company, id):
        return (
            "SELECT EXTFILENAME, NAME"
            + " FROM {}.{}.dbo.ZGEM_EMAILATTCHMNTS".format(server, company)
            + " WHERE EMAILID = {}".format(id)
            + " UNION"
            + " SELECT EXTFILENAME, ATTACHMENTNAME"
            + " FROM {}.{}.dbo.ZGEM_EMAILING E,".format(server, company)
            + " {}.system.dbo.ZGEM_DEFATTACHMNTS A".format(server)
            + " WHERE E.EMAILID = {}".format(id)
            + " AND E.TEMPLATE_EXEC = A.T$EXEC"
            + " AND E.TEMPLATE_NUM = A.NUM;"
        )

    def get_variables_statement(self, server, company, id):
        return (
            "SELECT VARIABLENAME, VALUE"
            + " FROM {}.{}.dbo.ZGEM_EMAILVARIABLES".format(server, company)
            + " WHERE EMAILID = {}".format(id)
        )
    
    def get_email_recipients(self, server, company, id):
        return (
            "SELECT EMAIL, TYPE"
            + " FROM {}.{}.dbo.ZGEM_EMAIL_RECIPIENT".format(server, company)
            + " WHERE EMAILID = {}".format(id)
        )
    
    def get_next_email_in_email_stack(self, server, company):
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
    
    def get_companies(self, server):
        return """
                SELECT DNAME
                FROM {}.system.dbo.ZGEM_ENVIRONMENT
                WHERE EMAIL_SYSTEM_ACTIVE = 'Y'
                AND DNAME <> ''
                """.format(server)
    
    def get_recipient_whitelist(self, server, company):
        return """
                SELECT VALUE
                FROM {}.{}.dbo.ZGEM_EMAIL_SETTINGS
                WHERE NAME = 'RECIPIENT_WHITELIST'
                """.format(server, company)
    
    def get_whitelisted_recipients(self, server, company):
        return """
                SELECT EMAIL
                FROM {}.{}.dbo.ZGEM_RECIPIENT_WL
                WHERE EMAIL <> ''
                """.format(server, company)
    
    def get_server_companies(self, server):
        """Get companies for a specific server"""
        config = Config(env='system', path=self.working_dir)
        cnxn = Config.zgem_cnxn(config, server.serverIP, server.serverCredentials)
        cursor = cnxn.cursor()
        companies = []

        cursor.execute(self.get_companies(server.server))
        results = cursor.fetchall()

        for row in results:
            companies.append(row[0])

        cursor.close()
        cnxn.close()
        server.companies = companies
        self.logger.debug(f"Server: {server.name}, Companies: {companies}")
        return companies
    
    def process_whitelist_filter(self, server, company, cnxn):
        """Check if recipient whitelist is enabled and get whitelisted emails"""
        cursor = cnxn.cursor()
        cursor.execute(f"USE [{company}];")
        
        recipient_whitelist_result = cursor.execute(self.get_recipient_whitelist(server.server, company))
        recipient_whitelist_result = cursor.fetchall()
        recipient_whitelist_filter = (
            recipient_whitelist_result and 
            len(recipient_whitelist_result) > 0 and 
            len(recipient_whitelist_result[0]) > 0 and 
            recipient_whitelist_result[0][0] == 'On'
        )
        self.logger.debug("Whitelisted filter: {}".format(recipient_whitelist_filter))
        whitelisted_emails = []
        if recipient_whitelist_filter:
            results = cursor.execute(self.get_whitelisted_recipients(server.server, company))
            results = cursor.fetchall()
            for row in results:
                whitelisted_emails.append(row[0])
            self.logger.debug("Whitelisted Emails: {}".format(whitelisted_emails))
        
        cursor.close()
        
        return recipient_whitelist_filter, whitelisted_emails
    
    def apply_whitelist_filter(self, email_recipients, whitelisted_emails):
        """Apply whitelist filter to email recipients"""
        email_recipients['To'] = [email for email in email_recipients['To'] if email in whitelisted_emails]
        email_recipients['Cc'] = [email for email in email_recipients['Cc'] if email in whitelisted_emails]
        email_recipients['Bcc'] = [email for email in email_recipients['Bcc'] if email in whitelisted_emails]
        
        # Ensure there's at least one recipient in 'To'
        if len(email_recipients['To']) <= 0:
            if len(email_recipients['Cc']) > 0:
                email_recipients['To'].append(email_recipients['Cc'].pop(0))
            elif len(email_recipients['Bcc']) > 0:
                email_recipients['To'].append(email_recipients['Bcc'].pop(0))
    
    def send_email(self, message, email_recipients, sender_email, server, id, email_type, subject):
        """Send email with fallback handling"""
        email_was_sent = 0
        try:
            with smtplib.SMTP(self.smtp_server, self.smtp_port) as smtp_server:
                smtp_server.starttls()
                smtp_server.login(self.priority_email, self.priority_password)
                all_recipients = email_recipients['To'] + email_recipients['Cc'] + email_recipients['Bcc']
                smtp_server.sendmail(self.priority_email, all_recipients, message.as_string())
                email_was_sent = 1
                self.logger.debug(f"Server: {server.name} ID: {id} Type: {email_type} From: {sender_email} To: {email_recipients} Subject: {subject}")
        except Exception as ex:
            if type(ex).__name__ == "SMTPDataError" and "SendAsDenied" in str(ex.args[1]):
                email_was_sent = self.handle_send_as_denied(message, email_recipients, sender_email, server, id, email_type, subject)
            else:
                error_msg = f"Server: {server.name} SMTP Email Connection Failed Error Type: {type(ex).__name__}, Arguments: {ex.args!r}"
                self.logger.critical(error_msg)
        
        return email_was_sent
    
    def handle_send_as_denied(self, message, email_recipients, sender_email, server, id, email_type, subject):
        """Handle SendAsDenied error by sending from priority email"""
        try:
            with smtplib.SMTP(self.smtp_server, self.smtp_port) as smtp_server:
                # Reset sender and recipient
                del message["From"]
                del message["To"]
                if "Cc" in message:
                    del message["Cc"]
                
                message["From"] = formataddr(('Clarkson Evans ERP System', self.priority_email))
                message['To'] = ', '.join(email_recipients['To']) if email_recipients['To'] else ''
                message['Cc'] = ', '.join(email_recipients['Cc']) if email_recipients['Cc'] else ''
                message['Cc'] = message['Cc'] + ", " + sender_email
                message.add_header('reply-to', sender_email)
                
                # Send email from Priority email to recipient and sender
                smtp_server.starttls()
                smtp_server.login(self.priority_email, self.priority_password)
                all_recipients = email_recipients['To'] + email_recipients['Cc'] + email_recipients['Bcc'] + [sender_email]
                smtp_server.sendmail(self.priority_email, all_recipients, message.as_string())
                
                self.logger.info(f"Server: {server.name} ID: {id} Type: {email_type} ({self.priority_email} can't send as {sender_email} but email was still sent) From: {self.priority_email} To: {email_recipients},{sender_email} Subject: {subject}")
                
                # Notify IT to set up send as permission
                self.send_it_notification(sender_email)
                return 1
                
        except Exception as ex1:
            error_msg = f"Server: {server.name} SMTP Email Connection Failed Error Type: {type(ex1).__name__}, Arguments: {ex1.args!r}"
            self.logger.critical(error_msg)
            return 0
    
    def send_it_notification(self, sender_email):
        """Send notification to IT about send-as permission setup"""
        try:
            with smtplib.SMTP(self.smtp_server, self.smtp_port) as smtp_server:
                send_as_message = MIMEMultipart('alternative')
                it_html = self.IT_send_as_setup_html.replace("<SENDFROM>", self.priority_email).replace("<SENDAS>", sender_email)
                text_part = MIMEText(html2text.html2text(it_html), 'plain')
                send_as_message.attach(text_part)
                html_part = MIMEText(it_html, 'html')
                send_as_message.attach(html_part)
                send_as_message['Subject'] = "Set Up Send As Permission"
                send_as_message['From'] = formataddr(('Clarkson Evans ERP System', self.priority_email))
                send_as_message['To'] = self.IT_email
                
                smtp_server.starttls()
                smtp_server.login(self.priority_email, self.priority_password)
                smtp_server.sendmail(self.priority_email, self.IT_email, send_as_message.as_string())
        except Exception as ex:
            self.logger.error(f"Failed to send IT notification: {ex}")
    
    def process_emails(self):
        """Main method to process emails for all servers"""
        for server in self.servers:
            self.process_server_emails(server)
    
    def process_server_emails(self, server):
        """Process emails for a specific server"""
        # Get companies for server
        companies = self.get_server_companies(server)
        
        for company in companies:
            self.process_company_emails(server, company)
    
    def process_company_emails(self, server, company):
        """Process emails for a specific company"""
        # Connect to database
        config = Config(env=company, path=self.working_dir)
        cnxn = Config.zgem_cnxn(config, server.serverIP, server.serverCredentials)
        cnxn2 = Config.zgem_cnxn(config, server.serverIP, server.serverCredentials)
        
        # Check recipient whitelist
        recipient_whitelist_filter, whitelisted_emails = self.process_whitelist_filter(server, company, cnxn)
        
        # Process emails in queue
        while True:
            cursor1 = cnxn.cursor()
            cursor1.execute(f"USE [{company}];")
            next_email_query = self.get_next_email_in_email_stack(server.server, company)
            emails = cursor1.execute(next_email_query)
            
            if emails.rowcount == 0:
                cursor1.close()
                break
            
            for email in emails:
                email_was_sent = self.process_single_email(
                    email, server, company, cursor1, 
                    recipient_whitelist_filter, whitelisted_emails
                )
                
                # Update database
                sent = "'Y'" if email_was_sent == 1 else "'N'"
                update = self.email_update_statement(server.server, company, email[0], sent)
                cursor2 = cnxn2.cursor()
                cursor2.execute(update)
                cursor2.commit()
                cursor2.close()
            
            cursor1.close()
        
        cnxn.close()
        cnxn2.close()
    
    def process_single_email(self, email, server, company, cursor, recipient_whitelist_filter, whitelisted_emails):
        """Process a single email"""
        # Extract email data
        id = email[0]
        sender_name = email[2]
        sender_email = email[3]
        email_recipients = {
            'To': [email[4]],
            'Cc': [],
            'Bcc': []
        }
        subject = email[5]
        exec = email[6]
        num = email[7]
        email_type = email[10]
        use_manual_emailbody = True if email[11].upper() == 'Y' else False
        
        # Get additional email recipients
        cursor.execute(self.get_email_recipients(server.server, company, id))
        email_recipient_results = cursor.fetchall()
        
        for recipient in email_recipient_results:
            email_address = recipient[0]
            recipient_type = recipient[1]
            if recipient_type in email_recipients:
                email_recipients[recipient_type].append(email_address)
        
        # Apply whitelist filter
        if recipient_whitelist_filter:
            self.apply_whitelist_filter(email_recipients, whitelisted_emails)
        
        if len(email_recipients['To']) <= 0:
            return 1  # Skip if no valid recipients
        
        # Build email message
        message = self.build_email_message(
            server, company, cursor, id, sender_name, sender_email, 
            email_recipients, subject, exec, num, use_manual_emailbody
        )
        
        # Send email
        return self.send_email(message, email_recipients, sender_email, server, id, email_type, subject)
    
    def build_email_message(self, server, company, cursor, id, sender_name, sender_email, 
                          email_recipients, subject, exec, num, use_manual_emailbody):
        """Build the complete email message"""
        # Get email template
        get_email_template = self.get_email_template_statement(server.server, company, exec, num, id, use_manual_emailbody)
        email_template = cursor.execute(get_email_template)
        html = self.default_html_email_body
        
        if email_template.rowcount != 0:
            html = ''
            prev_line = ''
            for line in email_template:
                gap = ' '
                if 'src' in prev_line or 'href' in prev_line or 'https' in prev_line:
                    if prev_line.count('"') % 2 != 0:
                        gap = ''
                html += gap + line[0]
                prev_line = line[0]
        
        # Replace variables in email template
        get_variables = self.get_variables_statement(server.server, company, id)
        variables = cursor.execute(get_variables)
        if variables.rowcount != 0:
            for variable in variables:
                html = html.replace(variable[0], variable[1])
        
        # Build email body
        msg_alternative = MIMEMultipart('alternative')
        text_part = MIMEText(html2text.html2text(html), 'plain')
        msg_alternative.attach(text_part)
        self.logger.debug(f"HTML message body: {html}")
        html_part = MIMEText(html, 'html')
        msg_alternative.attach(html_part)
        
        # Create main message
        message = MIMEMultipart('mixed')
        message['Subject'] = subject
        message['From'] = formataddr((sender_name, sender_email))
        message['To'] = ', '.join(email_recipients['To']) if email_recipients['To'] else ''
        message['Cc'] = ', '.join(email_recipients['Cc']) if email_recipients['Cc'] else ''
        message.attach(msg_alternative)
        
        # Add attachments
        self.add_attachments(message, server, company, cursor, id)
        
        return message
    
    def add_attachments(self, message, server, company, cursor, id):
        """Add attachments to email message"""
        get_attachments = self.get_attachment_statement(server.server, company, id)
        attachments = cursor.execute(get_attachments)
        
        if attachments.rowcount != 0:
            for attachment in attachments:
                if attachment[0]:
                    file_name = attachment[1]
                    url = server.domain + attachment[0].replace("../", "").replace('system/', '').replace('mail/', 'primail/')
                    self.logger.debug(f"Attachment: {attachment[0]}")
                    self.logger.debug(f"Attachment URL: {url}")
                    
                    path, file_extension = os.path.splitext(url)
                    response = requests.get(url)
                    new_file = str(Path(__file__).parent) + '/' + str(uuid.uuid4()) + file_extension
                    
                    with open(new_file, 'wb') as file:
                        file.write(response.content)
                    
                    with open(new_file, 'rb') as file:
                        message.attach(MIMEApplication(file.read(), Name=file_name + file_extension))
                    
                    os.remove(new_file)

    def main(self): 
        self.setup_servers()
        self.process_emails()        

if __name__ == '__main__':    
    win32serviceutil.HandleCommandLine(MySVC)
