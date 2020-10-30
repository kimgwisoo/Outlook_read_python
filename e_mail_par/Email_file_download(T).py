import email
import imaplib
import os

from email import encoders


class FetchEmail():
    ENCODING = 'UTF-8'

    connection = None
    error = None

    def __init__(self, mail_server, username, password):
        self.connection = imaplib.IMAP4_SSL('imap.gmail.com')
        self.connection.login('kimgs243@gmail.com', 'mypassword')
        self.connection.select(readonly=False)  # 메일을 읽은 상태로 표시
        self.save_attachment(self.connection)

    def close_connection(self):
        """ 
        Imap Server를 닫습니다. 
        """
        self.connection.close()

    def save_attachment(self, msg, download_folder="C:\\Users\\MCF\\Downloads"):
        """ 
        메시지가 주어지면 첨부파일을 지정된 폴더로 다운하게 됩니다. (default is /tmp)
        return : 첨부파일로 반환한다.
        Given a message, save its attachments to the specified 
        download folder (default is /tmp) 

        return: file path to attachment 
        """

        att_path = "No attachment found."
        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue

            filename = part.get_filename()
            att_path = os.path.join(download_folder, filename)

            if not os.path.isfile(att_path):
                fp = open(att_path, 'wb')
                fp.write(part.get_payload(decode=True))
                fp.close()
        return att_path

    def fetch_unread_messages(self):
        """ 
        읽지 않은 메일 검색
        """
        emails = []
        (result, messages) = self.connection.search(None, 'UnSeen')
        if result == "OK":
            for message in messages[0].split(' '):
                try:
                    ret, data = self.connection.fetch(message, '(RFC822)')
                except:
                    print("No new emails to read.")
                    self.close_connection()
                    exit()

                msg = email.message_from_string(data[0][1])
                if isinstance(msg, str) == False:
                    emails.append(msg)
                response, data = self.connection.store(
                    message, '+FLAGS', '\\Seen')

            return emails

        self.error = "Failed to retreive emails."
        return emails

    def parse_email_address(self, email_address):
        """ 
        Helper function to parse out the email address from the message 

        return: tuple (name, address). Eg. ('John Doe', 'jdoe@example.com') 
        """
        return email.utils.parseaddr(email_address)


email = FetchEmail
email.save_attachment()
