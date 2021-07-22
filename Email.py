
import smtplib
from email.mime.text import MIMEText

smtp = smtplib.SMTP('smtp.live.com', 587)
smtp.ehlo()  # say Hello
smtp.starttls()  # TLS 사용시 필요
smtp.login('ksj10111011@gmail.com', 'utzexawjxjtbcbnu')

msg = MIMEText('본문 테스트 메시지')
msg['Subject'] = '테스트'
msg['To'] = 'kim@naver.com'
smtp.sendmail('ksj10111011@gmail.com', 'ksj10111011@gmail.com', msg.as_string())

smtp.quit()