import smtplib
import threading
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart



class EmailNotifier:
    def __init__(self, sender_email, sender_password, receiver_email):

        self.sender_email = "vlbi.alarm@gmail.com"
        self.sender_password = "kqgxndxnbsqvahko"
        self.receiver_email = "lan.ttn124@gmail.com"
        self.smtp_server = "smtp.gmail.com"
        self.smtp_port = 465  # Port cho SSL

    def send_alert_thread(self, subject, message_body):
        """이 함수는 이메일 전송 프로세스를 실행(별도의 스레드에서 실행됨).."""
        try:
            msg = MIMEMultipart()
            msg['From'] = self.sender_email
            msg['To'] = self.receiver_email
            msg['Subject'] = subject

            msg.attach(MIMEText(message_body, 'plain'))

            # Gmail 서버에 연결
            with smtplib.SMTP_SSL(self.smtp_server, self.smtp_port) as server:
                server.login(self.sender_email, self.sender_password)
                server.send_message(msg)

            print(f"-> Email sent: {subject}")

        except Exception as e:
            print(f"Error sending email: {e}")

    def send_warning(self, device_time, device, alarm_level, msg):
        subject = f"[VLBI SYSTEM WARNING]"
        body = (f"[{device_time}] [{alarm_level}] {device}_{msg}")

        # GUI 충돌을 방지하려면 이메일 전송 작업을 별도의 스레드에서 실행
        email_thread = threading.Thread(target=self.send_alert_thread, args=(subject, body))
        email_thread.start()

    def send_caution(self, device_time, device, alarm_level, msg ):
        subject = f"[VLBI SYSTEM CAUTION]"
        body = (f"[{device_time}] [{alarm_level}] {device} {msg}")

        email_thread = threading.Thread(target=self.send_alert_thread, args=(subject, body))
        email_thread.start()

    def send_recovery_email(self, device, col):
        """Mframe_left에서 호출할 때 발생하는 오류를 방지하려면 이 함수를 추가"""
        subject = f"[VLBI RECOVERY] {device} - {col}"
        body = f"{device} - {col} 정상으로 돌아왔습니다."
        email_thread = threading.Thread(target=self.send_alert_thread, args=(subject, body))
        email_thread.start()