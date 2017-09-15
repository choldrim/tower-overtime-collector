import imghdr
import os
import smtplib

from configparser import ConfigParser
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

EMAIL_CONF = 'config.ini'

FOOTER = """
<br>
<br>
<hr>
<footer>
  <p>Posted by: <a href="https://ci.deepin.io">deepin AutomatedScript ci</a></p>
  <p>Contact: <a href="mailto:tangcaijun@linuxdeepin.com">tangcaijun@linuxdeepin.com</a></p>
</footer>
"""


class Email:
    def __init__(self):
        (server, username, passwd) = self.__get_user_info()
        self.smtp = self.__smtp(server, username, passwd)
        self.sender = username


    def __smtp(self, server, username, passwd):
        try:
            smtp = smtplib.SMTP()
            smtp.connect(server)
            smtp.login(username, passwd)
            smtp.helo()
            return smtp
        except Exception as e:
            print("E: connect mailing service fail")
            raise e


    def __get_user_info(self):
        config = ConfigParser()
        config.read(EMAIL_CONF)
        server = config["EMAIL"]["SMTPServer"]
        username = config["EMAIL"]["UserName"]
        passwd = config["EMAIL"]["UserPWD"]
        return server, username, passwd


    def send(self, receiver, subject, content, CC="", files=[], auto_close=False, use_footer=False):
        """
        param examples:
            receiver: a@domain.com,b@domain.com,c@domain.com
            subject: subject_str
            content: <h1>hello world</h1>
            CC: "h@domain.com,i@domain.com"
            files: ["file_1.jpg", "/home/deepin/file_2.txt", "/tmp/file_3.doc"]
        """
        root = MIMEMultipart()
        root["Subject"] = subject
        root["From"] = self.sender
        root["To"] = receiver
        if CC:
            root["Cc"] = CC

        if use_footer:
            content_part = MIMEText(content + FOOTER, "html", "utf-8")
        else:
            content_part = MIMEText(content, "html", "utf-8")

        root.attach(content_part)

        if CC:
            all_receivers = receiver.split(",") + CC.split(",")
        else:
            all_receivers = receiver.split(",")

        for send_file in files:
            att = MIMEText(open(send_file, "rb").read(), "base64", "UTF-8")
            att["Content-Type"] = "application/octet-stream"
            att["Content-Disposition"] = 'attachment; filename="%s"' % os.path.basename(send_file)
            root.attach(att)

        self.smtp.sendmail(self.sender, all_receivers, root.as_string())

        if auto_close:
            self.close()


    def close(self):
        self.smtp.quit()

