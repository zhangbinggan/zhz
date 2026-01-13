import smtplib
import logging
import os
from email.mime.text import MIMEText
from email.header import Header


# 推送到邮箱
def feishu(DD_BOT_TOKEN, DD_BOT_SECRET, text, desp):
    """
    发送邮箱通知

    Args:
        DD_BOT_TOKEN: 钉钉令牌（未使用，保持与钉钉函数签名一致）
        DD_BOT_SECRET: 钉钉密钥（未使用，保持与钉钉函数签名一致）
        text: 消息标题
        desp: 消息内容

    Returns:
        dict: 发送结果
    """
    # 配置信息
    sender_email = "166767710@qq.com"
    receiver_email = os.environ.get("FEISHU_BOT_SECRET")
    smtp_server = "smtp.qq.com"
    smtp_port = 465  # 使用SSL加密端口
    password = "teekuuhqnbrncbag"  # QQ邮箱授权码
    
    # 检查收件人邮箱是否配置
    if not receiver_email:
        logging.error("收件人邮箱未配置，请在环境变量中设置FEISHU_BOT_SECRET")
        return {"success": False, "message": "收件人邮箱未配置"}
    
    # 创建邮件内容，格式与钉钉相同：text\ndesp
    email_content = f"{text}\n{desp}"
    message = MIMEText(email_content, 'plain', 'utf-8')
    message['From'] = Header(sender_email)
    message['To'] = Header(receiver_email)
    message['Subject'] = Header(text, 'utf-8')  # 使用text作为邮件主题
    
    try:
        # 连接SMTP服务器
        server = smtplib.SMTP_SSL(smtp_server, smtp_port)
        logging.info(f"成功连接到SMTP服务器: {smtp_server}:{smtp_port}")
        
        # 登录邮箱
        server.login(sender_email, password)
        logging.info(f"邮箱登录成功: {sender_email}")
        
        # 发送邮件
        server.sendmail(sender_email, receiver_email, message.as_string())
        logging.info(f"邮件发送成功🎉\n收件人: {receiver_email}\n主题: {text}")
        
        # 关闭连接
        server.quit()
        return {"success": True, "message": "邮件发送成功"}
        
    except Exception as e:
        logging.error(f"邮件发送失败😞\n收件人: {receiver_email}\n主题: {text}\n错误信息: {str(e)}")
        return {"success": False, "message": f"邮件发送失败: {str(e)}"}
