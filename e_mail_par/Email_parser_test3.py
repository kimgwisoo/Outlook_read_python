import imaplib
import email

# 문자열의 인코딩 정보추출 후, 문자열, 인코딩 얻기


def findEncodingInfo(txt):
    info = email.header.decode_header(txt)
    s, encoding = info[0]
    return s, encoding


# 메일 서버 로그인
imap = imaplib.IMAP4_SSL('imap.gmail.com')
id = 'kimgs243@gmail.com'
pw = 'lizrrkqakqbjshkv'
imap.login(id, pw)

# 받은 편지함
imap.select('inbox')

# 받은 편지함 모든 메일 검색
resp, data = imap.uid('search', None, 'All')

# 여러 메일 읽기 (반복)
all_email = data[0].split()

for mail in all_email:

    # fetch 명령을 통해서 메일 가져오기 (RFC822 Protocol)
    result, data = imap.uid('fetch', mail, '(RFC822)')

    # 사람이 읽기 힘든 Raw 메세지 (byte)
    raw_email = data[0][1]

    # 메시지 처리(email 모듈 활용)
    email_message = email.message_from_bytes(raw_email)

    # 이메일 정보 keys
    # print(email_message.keys())
    print('FROM:', email_message['From'])
    print('SENDER:', email_message['Sender'])
    print('TO:', email_message['To'])
    print('DATE:', email_message['Date'])

    b, encode = findEncodingInfo(email_message['Subject'])
    print('SUBJECT:', str(b, encode))

    # 이메일 본문 내용 확인
    print('[CONTENT]')
    print('='*80)
    if email_message.is_multipart():
        for part in email_message.get_payload():
            bytes = part.get_payload(decode=True)
            encode = part.get_content_charset()
            print(str(bytes, encode))
    print('='*80)

imap.close()
imap.logout()
