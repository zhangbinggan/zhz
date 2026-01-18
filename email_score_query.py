import smtplib
import imaplib
import email
import json
import logging
import os
import re
from email.mime.text import MIMEText
from email.header import Header
from email.utils import parseaddr
import datetime
import time
from dotenv import load_dotenv

# 加载环境变量
load_dotenv()

# 获取配置的学号
CONFIG_USER_ACCOUNT = os.getenv("USER_ACCOUNT", "")

# 设置日志配置
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler("email_query.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)

# 邮箱配置信息
EMAIL_CONFIG = {
    "email": "166767710@qq.com",
    "smtp_server": "smtp.qq.com",
    "smtp_port": 465,
    "imap_server": "imap.qq.com",
    "imap_port": 993,
    "password": "teekuuhqnbrncbag"  # QQ邮箱授权码
}

# 成绩文件路径
SCORES_FILE = "scores.json"

def load_scores():
    """
    从scores.json文件中加载成绩数据
    """
    try:
        if not os.path.exists(SCORES_FILE):
            logging.error(f"成绩文件不存在: {SCORES_FILE}")
            return {}
            
        with open(SCORES_FILE, "r", encoding="utf-8") as f:
            scores_data = json.load(f)
            
        # 转换为字典格式以便快速查询 {科目名称: 成绩}
        scores_dict = {}
        for item in scores_data:
            if len(item) >= 2:
                subject = item[0].strip()
                score = item[1].strip()
                scores_dict[subject] = score
                
        logging.info(f"成功加载 {len(scores_dict)} 门课程成绩")
        return scores_dict
        
    except Exception as e:
        logging.error(f"加载成绩文件失败: {str(e)}")
        return {}

def search_query_emails():
    """
    搜索邮箱中格式为"查询成绩科目"的邮件
    返回: [(发件人邮箱, 发件人名称, 科目名称), ...]
    """
    query_emails = []
    
    try:
        # 连接IMAP服务器
        mail = imaplib.IMAP4_SSL(EMAIL_CONFIG['imap_server'], EMAIL_CONFIG['imap_port'])
        logging.info(f"✓ 成功连接到IMAP服务器: {EMAIL_CONFIG['imap_server']}:{EMAIL_CONFIG['imap_port']}")
        
        # 登录邮箱
        mail.login(EMAIL_CONFIG['email'], EMAIL_CONFIG['password'])
        logging.info(f"✓ 邮箱登录成功: {EMAIL_CONFIG['email']}")
        
        # 选择收件箱
        mail.select('inbox')
        logging.info(f"✓ 成功选择收件箱")
        
        # 搜索未读邮件
        status, messages = mail.search(None, 'UNSEEN')
        
        if status != 'OK':
            logging.error("✗ 搜索邮件失败")
            return query_emails
            
        # 获取邮件ID列表
        email_ids = messages[0].split()
        total_emails = len(email_ids)
        logging.info(f"✓ 找到 {total_emails} 封未读邮件")
        
        if total_emails == 0:
            return query_emails
            
        # 只处理最新的两条未读邮件
        emails_to_process = email_ids[-2:] if total_emails >= 2 else email_ids
        logging.info(f"✓ 仅处理最新的 {len(emails_to_process)} 封未读邮件")
            
        # 逐个检查邮件
        for email_id in emails_to_process:
            # 获取邮件内容
            status, msg_data = mail.fetch(email_id, '(RFC822)')
            
            if status != 'OK':
                logging.error(f"✗ 获取邮件失败: {email_id}")
                continue
                
            # 解析邮件
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    
                    # 获取发件人信息
                    sender = msg['From']
                    if sender:
                        sender_name, sender_addr = parseaddr(sender)
                        if sender_name:
                            try:
                                sender_name = email.header.decode_header(sender_name)[0][0]
                                if isinstance(sender_name, bytes):
                                    sender_name = sender_name.decode('utf-8')
                            except:
                                sender_name = "未知"
                    
                    # 获取邮件主题
                    subject = msg['Subject']
                    if subject:
                        try:
                            subject = email.header.decode_header(subject)[0][0]
                            if isinstance(subject, bytes):
                                subject = subject.decode('utf-8')
                        except:
                            subject = "无主题"
                    
                    # 获取邮件内容
                    body = ""
                    if msg.is_multipart():
                        for part in msg.walk():
                            content_type = part.get_content_type()
                            content_disposition = str(part.get("Content-Disposition"))
                            
                            if content_type == 'text/plain' and 'attachment' not in content_disposition:
                                try:
                                    body = part.get_payload(decode=True).decode('utf-8')
                                except:
                                    try:
                                        body = part.get_payload(decode=True).decode('gbk')
                                    except:
                                        body = "无法解析的内容"
                                break
                    else:
                        try:
                            body = msg.get_payload(decode=True).decode('utf-8')
                        except:
                            try:
                                body = msg.get_payload(decode=True).decode('gbk')
                            except:
                                body = "无法解析的内容"
                    
                    # 确保主题和内容都是字符串
                    if not isinstance(subject, str):
                        subject = str(subject)
                    if not isinstance(body, str):
                        body = str(body)
                    
                    # 检查邮件是否包含"查询成绩"或"成绩查询"关键字
                    full_content = f"{subject}\n{body}"
                    # 使用小写进行匹配
                    full_content_lower = full_content.lower()
                    if "查询成绩" in full_content_lower or "成绩查询" in full_content_lower:
                        # 提取学号和科目名称
                        # 支持的格式："查询成绩学号科目" 或 "成绩查询学号科目" 或 "查询成绩:学号科目" 或 "成绩查询:学号科目"或带空格的格式
                        pattern = r"(?:查询成绩|成绩查询)[\s:：]*([0-9]*)[\s:：]*(.+)" 
                        match = re.search(pattern, full_content_lower, re.IGNORECASE)
                        if match:
                            user_account = match.group(1).strip()
                            subject_name = match.group(2).strip()
                            
                            # 去除多余的标点符号和空格
                            subject_name = re.sub(r'[^\u4e00-\u9fa5a-zA-Z0-9\s]', '', subject_name)
                            # 去除所有空格（包括Unicode空格）
                            subject_name = re.sub(r'[\s\u00A0\u2000-\u200F\u2028-\u202F\u205F\u3000]+', '', subject_name)
                            
                            if subject_name:
                                # 如果没有提供学号，使用默认学号
                                if not user_account:
                                    user_account = CONFIG_USER_ACCOUNT
                                    
                                query_emails.append((sender_addr, sender_name, user_account, subject_name))
                                logging.info(f"✓ 发现成绩查询请求: 发件人={sender_addr}, 学号={user_account}, 科目={subject_name}")
                                
                                # 将邮件标记为已读
                                mail.store(email_id, '+FLAGS', r'\Seen')
        
        # 关闭连接
        mail.close()
        mail.logout()
        logging.info(f"✓ 邮箱连接已关闭")
        
    except Exception as e:
        logging.error(f"✗ 搜索查询邮件失败: {str(e)}")
        
    return query_emails

def send_score_email(recipient_email, recipient_name, user_account, subject_name, matching_courses, scores_dict):
    """
    发送成绩查询结果邮件
    """
    try:
        # 连接SMTP服务器
        server = smtplib.SMTP_SSL(EMAIL_CONFIG['smtp_server'], EMAIL_CONFIG['smtp_port'])
        logging.info(f"✓ 成功连接到SMTP服务器: {EMAIL_CONFIG['smtp_server']}:{EMAIL_CONFIG['smtp_port']}")
        
        # 登录邮箱
        server.login(EMAIL_CONFIG['email'], EMAIL_CONFIG['password'])
        logging.info(f"✓ 邮箱登录成功: {EMAIL_CONFIG['email']}")
        
        # 验证学号是否匹配
        if user_account != CONFIG_USER_ACCOUNT:
            logging.warning(f"✗ 学号不匹配: 请求学号={user_account}, 配置学号={CONFIG_USER_ACCOUNT}")
            
            email_subject = f"成绩查询结果 - 学号不匹配"
            email_content = f"尊敬的{recipient_name}：\n\n"
            email_content += f"您查询的学号{user_account}与系统配置学号不匹配，请检查后重试。\n\n"
            email_content += f"查询时间：{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            email_content += f"成绩监控系统"
        else:
            # 构建邮件内容
            if matching_courses:
                email_subject = f"成绩查询结果 - {subject_name}"
                
                # 构建邮件内容
                email_content = f"尊敬的{recipient_name}：\n\n"
                email_content += f"您查询的学号{user_account}的{subject_name}相关成绩如下：\n\n"
                
                for course, course_score in matching_courses:
                    email_content += f"• {course}：{course_score}\n"
                
                email_content += f"共有 {len(matching_courses)} 门相关课程\n\n"
                email_content += f"查询时间：{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                email_content += f"成绩监控系统"
            else:
                email_subject = f"成绩查询结果 - {subject_name}"
                email_content = f"尊敬的{recipient_name}：\n\n"
                email_content += f"未找到与 '{subject_name}' 相关的课程成绩。\n\n"
                email_content += f"查询时间：{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                email_content += f"成绩监控系统"
        
        # 创建邮件对象
        message = MIMEText(email_content, 'plain', 'utf-8')
        message['From'] = Header(EMAIL_CONFIG['email'])
        message['To'] = Header(recipient_email)
        message['Subject'] = Header(email_subject, 'utf-8')
        
        # 发送邮件
        server.sendmail(EMAIL_CONFIG['email'], recipient_email, message.as_string())
        logging.info(f"✓ 成绩查询结果已发送至: {recipient_email}")
        
        # 关闭连接
        server.quit()
        logging.info(f"✓ SMTP连接已关闭")
        
        return True
        
    except Exception as e:
        logging.error(f"✗ 发送成绩邮件失败: {str(e)}")
        return False

def process_score_queries():
    """
    处理所有成绩查询请求
    """
    try:
        logging.info("\n" + "="*60)
        logging.info("开始处理成绩查询请求")
        logging.info("="*60)
        
        # 加载成绩数据
        scores_dict = load_scores()
        if not scores_dict:
            logging.error("✗ 没有成绩数据可查询")
            return
            
        # 搜索查询邮件
        logging.info("✓ 开始搜索查询邮件")
        query_emails = search_query_emails()
        logging.info(f"✓ 搜索到 {len(query_emails)} 个查询请求")
    
        if not query_emails:
            logging.info("✓ 没有新的成绩查询请求")
            return
    
        logging.info(f"✓ 共发现 {len(query_emails)} 个成绩查询请求")
    
        # 处理每个查询请求
        for recipient_email, recipient_name, user_account, subject_name in query_emails:
            # 查找成绩（支持模糊匹配）
            logging.info(f"✓ 处理查询: 学号={user_account}, 科目={subject_name}")
            matching_courses = []
            
            for course, course_score in scores_dict.items():
                if subject_name in course or course in subject_name:
                    matching_courses.append((course, course_score))
    
            if matching_courses:
                logging.info(f"✓ 找到 {len(matching_courses)} 门相关课程: {[c[0] for c in matching_courses]}")
            else:
                logging.info(f"✗ 未找到与 '{subject_name}' 相关的课程")
            
            # 发送查询结果邮件
            logging.info(f"✓ 准备发送结果到: {recipient_email}")
            send_score_email(recipient_email, recipient_name, user_account, subject_name, matching_courses, scores_dict)
    
        logging.info("✓ 所有成绩查询请求已处理完成")
    except Exception as e:
        logging.error(f"✗ 处理成绩查询请求时发生错误: {str(e)}")

def run_email_score_query():
    """
    主函数，运行邮件成绩查询服务
    """
    logging.info("\n" + "="*60)
    logging.info("邮件成绩查询服务已启动")
    logging.info("="*60)
    logging.info("服务功能：")
    logging.info("1. 定期检查邮箱中的成绩查询请求")
    logging.info("2. 格式要求：邮件主题或内容包含'查询成绩[学号]科目名称'")
    logging.info("   - 示例：查询成绩202311103085软件测试基础")
    logging.info("   - 学号可选，未提供时使用默认配置的学号")
    logging.info("3. 自动从成绩文件中查找相关成绩")
    logging.info("4. 仅当查询学号与配置学号匹配时返回成绩")
    logging.info("5. 将查询结果发送回发件人邮箱")
    logging.info("="*60)
    
    try:
        while True:
            process_score_queries()
            
            # 等待一段时间后再次检查
            wait_time = 60  # 每60秒检查一次
            logging.info(f"✓ 将在 {wait_time} 秒后再次检查邮箱...")
            time.sleep(wait_time)
            
    except KeyboardInterrupt:
        logging.info("\n✓ 程序已中断")
    except Exception as e:
        logging.error(f"✗ 程序运行错误: {str(e)}")

if __name__ == "__main__":
    run_email_score_query()