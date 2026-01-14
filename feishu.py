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
    # 主邮箱配置信息
    primary_config = {
        "sender_email": "166767710@qq.com",
        "smtp_server": "smtp.qq.com",
        "smtp_port": 465,  # 使用SSL加密端口
        "password": "teekuuhqnbrncbag"  # QQ邮箱授权码
    }
    
    # 备用邮箱配置信息
    backup_config = {
        "sender_email": "166767710@qq.com",
        "smtp_server": "smtp.qq.com",
        "smtp_port": 465,  # 使用SSL加密端口
        "password": "lufbetebyleobhcb"  # QQ邮箱授权码
    }
    
    receiver_emails_env = os.environ.get("FEISHU_BOT_SECRET")
    
    # 检查收件人邮箱是否配置
    if not receiver_emails_env:
        logging.error("收件人邮箱未配置，请在环境变量中设置FEISHU_BOT_SECRET")
        return {"success": False, "message": "收件人邮箱未配置"}
    
    # 将逗号分隔的邮箱列表转换为Python列表
    receiver_emails = [email.strip() for email in receiver_emails_env.split(',')]
    
    # 创建邮件内容，格式与钉钉相同：text\ndesp
    email_content = f"{text}\n{desp}"
    
    def send_emails(config):
        """
        使用指定配置发送邮件
        
        Args:
            config: 邮箱配置字典
            
        Returns:
            tuple: (成功数量, 失败数量, 失败的邮箱列表)
        """
        sender_email = config["sender_email"]
        smtp_server = config["smtp_server"]
        smtp_port = config["smtp_port"]
        password = config["password"]
        
        try:
            # 连接SMTP服务器
            server = smtplib.SMTP_SSL(smtp_server, smtp_port)
            logging.info(f"成功连接到SMTP服务器: {smtp_server}:{smtp_port}")
            
            # 登录邮箱
            server.login(sender_email, password)
            logging.info(f"邮箱登录成功: {sender_email}")
            
            success_count = 0
            failed_emails = []
            
            # 对每个收件人单独发送邮件
            for receiver_email in receiver_emails:
                try:
                    # 为每个收件人创建独立的邮件对象
                    single_message = MIMEText(email_content, 'plain', 'utf-8')
                    single_message['From'] = Header(sender_email)
                    single_message['To'] = Header(receiver_email)
                    single_message['Subject'] = Header(text, 'utf-8')
                    
                    # 发送邮件给单个收件人
                    server.sendmail(sender_email, receiver_email, single_message.as_string())
                    logging.info(f"邮件发送成功🎉\n发件人: {sender_email}\n收件人: {receiver_email}\n主题: {text}")
                    success_count += 1
                except Exception as single_e:
                    logging.error(f"邮件发送失败😞\n发件人: {sender_email}\n收件人: {receiver_email}\n主题: {text}\n错误信息: {str(single_e)}")
                    failed_emails.append(receiver_email)
            
            # 关闭连接
            server.quit()
            
            return success_count, len(failed_emails), failed_emails
            
        except Exception as e:
            logging.error(f"SMTP操作失败😞\n发件人: {sender_email}\n错误信息: {str(e)}")
            return 0, len(receiver_emails), receiver_emails
    
    # 使用主邮箱发送邮件
    primary_success, primary_fail, failed_emails = send_emails(primary_config)
    
    total_success = primary_success
    total_fail = primary_fail
    
    # 如果有失败的邮箱，使用备用邮箱重新发送
    if failed_emails:
        logging.info(f"使用备用邮箱重新发送 {len(failed_emails)} 封失败的邮件")
        
        # 更新收件人列表为失败的邮箱列表
        receiver_emails = failed_emails
        
        # 使用备用邮箱发送
        backup_success, backup_fail, final_failed_emails = send_emails(backup_config)
        
        total_success += backup_success
        total_fail = backup_fail
    
    # 返回整体发送结果
    if total_fail == 0:
        return {"success": True, "message": f"邮件发送成功，共发送 {total_success} 封邮件"}
    elif total_success == 0:
        return {"success": False, "message": f"所有邮件发送失败，共失败 {total_fail} 封邮件"}
    else:
        return {"success": False, "message": f"部分邮件发送失败，成功 {total_success} 封，失败 {total_fail} 封"}
