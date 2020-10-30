import os
import email
import imaplib
import email.header
import configparser
import base64


# 문자열의 인코딩 정보 추출 후, 문자열, 인코딩 얻기


def find_encoding_info(txt):
    info = email.header.decode_header(txt)
    s, encoding = info[0]
    return s, encoding


# Email 설정정보 불러오기
config = configparser.ConfigParser()
config.read('config.ini')

# gmail imap 세션 생성
imap_url = imaplib.IMAP4_SSL('imap.gmail.com')

# 로그인
imap_url.login(config['Gmail']['id'], config['Gmail']['pw'])

# 받은편지함
imap_url.select('INBOX')

# 받은 편지함 내 모든 메일 검색
result, data = imap_url.search(None, 'ALL')

# 첨부파일 경로
detach_dir = 'D:\개인\Work-2\Download'

# 여러 메일 읽기
all_email = data[0].split()

for mail in all_email:
    result, data = imap_url.fetch(mail, '(RFC822)')
    raw_email = data[0][1]
    raw_email_string = raw_email.decode('utf-8')
    email_message = email.message_from_string(raw_email_string)

    # 메일 정보
    print('From: ', email_message['From'])
    print('Sender: ', email_message['Sender'])
    print('To: ', email_message['To'])
    print('Date: ', email_message['Date'])

    subject, encode = find_encoding_info(email_message['Subject'])
    print('Subject', subject)

    message = ''

    print('[Message]')
    # 메일 본문 확인
    if email_message.is_multipart():
        for part in email_message.get_payload():
            if part.get_content_type() == 'text/plain':
                bytes = part.get_payload(decode=True)
                encode = part.get_content_charset()
                message = message + str(bytes, encode)
    else:
        if email_message.get_content_type() == 'text/plain':
            bytes = email_message.get_payload(decode=True)
            encode = email_message.get_content_charset()
            message = str(bytes, encode)
    print(message)

    # 첨부파일 존재 시 다운로드
    for part in email_message.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        file_name = part.get_filename()

        # file_name = "".join(
        #     i for i in file_name if i not in "\/:*?<>|")

        # file_name = file_name.replace('UTF-8B', '')
        # file_name = file_name.replace('utf-8B', '')
        # print(file_name)

        if bool(file_name):
            file_path = os.path.join(detach_dir, file_name)
            if not os.path.isfile(file_path):
                fp = open(file_path, 'wb')
                fp.write(part.get_payload(decode=True))
                fp.close()
    else:
        continue

imap_url.close()
imap_url.logout()
