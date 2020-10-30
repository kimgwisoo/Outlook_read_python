import imaplib
import email

# Connect to an IMAP server


def connect(server, user, password):
    m = imaplib.IMAP4_SSL('imap.gmail.com')
    m.login(user, password)
    m.select()
    return m


# 해당 이메일에 대한 모든 첨부 파일을 다운로드

def downloaAttachmentsInEmail(m, emailid, outputdir):
    resp, data = m.fetch(emailid, "(BODY.PEEK[])")
    email_body = data[0][1].decode('UTF-8')
    mail = email.message_from_string(email_body)
    if mail.get_content_maintype() != 'multipart':
        return
    for part in mail.walk():
        if part.get_content_maintype() != 'multipart' and part.get('Content-Disposition') is not None:
            open(outputdir + '/' + part.get_filename(),
                 'wb').write(part.get_payload(decode=True))


# 받은 편지함의 모든 이메일에 대한 모든 첨부 파일을 다운로드합니다.
def downloadAllAttachmentsInInbox(server, user, password, outputdir):
    m = connect(server, user, password)
    resp, items = m.search(None, "(ALL)")
    items = items[0].split()
    for emailid in items:
        downloaAttachmentsInEmail(m, emailid, outputdir)


# connect(imaplib.IMAP4_SSL('imap.gmail.com'),
#         'kimgs243@gmail.com', 'lizrrkqakqbjshkv')
downloaAttachmentsInEmail('kimgs243@gmail.com',
                          'lizrrkqakqbjshkv', 'C:/Users/MCF/Downloads')
