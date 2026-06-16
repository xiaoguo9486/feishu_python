import smtplib
from email.mime.text import MIMEText
from email.header import Header
from datetime import datetime   # 新增导入


def send_mail(smtp_server, smtp_port, user, password, to_addrs, subject, body):
    if isinstance(to_addrs, str):
        to_addrs = [addr.strip() for addr in to_addrs.split(',') if addr.strip()]
    if not to_addrs:
        return False, "收件人地址为空"

    try:
        msg = MIMEText(body, 'plain', 'utf-8')
        msg['From'] = user
        msg['To'] = ', '.join(to_addrs)
        msg['Subject'] = subject          # 直接赋值，无需 Header

        port = int(smtp_port)
        if port == 465:
            server = smtplib.SMTP_SSL(smtp_server, port, timeout=15)
        else:
            server = smtplib.SMTP(smtp_server, port, timeout=15)
            server.ehlo()
            if port == 587:
                server.starttls()
                server.ehlo()

        server.login(user, password)
        server.sendmail(user, to_addrs, msg.as_string().encode('utf-8'))
        server.quit()
        return True, f"邮件已发送至 {', '.join(to_addrs)}"
    except Exception as e:
        return False, f"邮件发送失败: {e}"


if __name__ == "__main__":
    print("=== 邮件发送功能测试 ===")
    # ========== QQ邮箱配置 ==========
    TEST_SMTP_SERVER = "smtp.qq.com"
    TEST_SMTP_PORT = 587
    TEST_USER = "2604182970@qq.com"  # 例如 12345678@qq.com
    TEST_PASSWORD = "tjeqnwmhwyiyecbg"  # 注意是授权码，不是QQ密码
    TEST_TO = "330951244@qq.com,xiaoguo9486@163.com"  # 收件人，可以是任意邮箱
    TEST_SUBJECT = "巡检照片整理 - 测试邮件"
    TEST_BODY = f"自动化测试邮件，发送时间：{datetime.now()}"
    # ================================

    success, msg = send_mail(
        TEST_SMTP_SERVER, TEST_SMTP_PORT,
        TEST_USER, TEST_PASSWORD,
        TEST_TO,
        TEST_SUBJECT, TEST_BODY
    )
    if success:
        print("✓ 成功:", msg)
    else:
        print("✗ 失败:", msg)