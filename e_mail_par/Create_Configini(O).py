import configparser

# ConfigParser 객체 생성
config = configparser.ConfigParser()

# 세션 생성
config['Gmail'] = {}

# option, value 생성 방법 1
config['Gmail']['ID'] = 'kimgs243@gmail.com'
config['Gmail']['PW'] = 'mypassword'

# option, value 생성 방법 2
Gmail = config['Gmail']
Gmail['Subject'] = '테스트'
Gmail['Body'] = '테스트입니다.'

# option, value 생성 방법 3
# config['Mail_info'] = {'ID': 'test', 'PW': 'test1234', 'Subject': '테스트', 'Body': '테스트입니다.'}

# config.ini 파일 생성
with open('config.ini', 'wt', encoding='UTF-8') as conf_file:
    config.write(conf_file)
