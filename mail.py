# -*- encoding: utf-8 -*-
import imaplib
from datetime import datetime
from email import policy
from email.parser import BytesParser
import re
import requests


class OutlookEmail:
    def __init__(self, email, client_id, token, imap_server='outlook.office365.com'):
        self.email = email  # 用户的邮箱地址
        self.imap_server = imap_server  # IMAP服务器地址，默认为Outlook的服务器
        self.mail = self.login(email, client_id, token)  # 登录邮箱，并保存登录状态

    @staticmethod
    def get_access_token(client_id, refresh_token):
        data = {
            'client_id': client_id,
            'grant_type': 'refresh_token',
            'refresh_token': refresh_token
        }
        ret = requests.post('https://login.live.com/oauth20_token.srf', data=data)
        return ret.json()['access_token']

    @staticmethod
    def generate_auth_string(user, token):
        auth_string = f"user={user}\1auth=Bearer {token}\1\1"
        return auth_string

    def connect_imap(self, email, access_token):
        imap = imaplib.IMAP4_SSL(self.imap_server)
        auth_string = self.generate_auth_string(email, access_token)
        imap.authenticate("XOAUTH2", lambda x: auth_string)
        return imap

    # 登录邮箱的方法
    def login(self, email, client_id, token):
        acc_token = self.get_access_token(client_id, token)
        mail = self.connect_imap(email, acc_token)
        mail.select('inbox')  # 选择"inbox"文件夹
        return mail  # 返回mail对象，表示已登录的状态

    # 静态方法，用于获取邮件ID列表
    @staticmethod
    def get_email_id_list(email_ids):
        email_id = email_ids[0].split()
        if len(email_id) != 0:
            return email_ids[0].split()
        return []

    # 获取单个邮件的详细信息
    def get_email_info(self, email_id):
        item = {}  # 创建一个空字典用于存储邮件信息
        # 获取指定ID的邮件数据
        status, email_data = self.mail.fetch(email_id, '(RFC822)')
        if status != 'OK':
            print("Error get email info")
            exit()
        # 解析邮件数据
        email_message = BytesParser(policy=policy.default).parsebytes(email_data[0][1])
        # 提取邮件的主题、发件人和日期
        item['subject'] = email_message['subject']
        item['from'] = email_message['from']
        item['date'] = email_message['date']

        # 检查邮件是否是多部分的
        if email_message.is_multipart():
            parts = email_message.get_payload()  # 获取邮件的所有部分
            # 提取纯文本部分
            for part in parts:
                if part.get_content_type() == 'text/plain':
                    try:
                        body = part.get_payload(decode=True).decode(part.get_content_charset())
                    except:
                        body = part.get_payload()
                    item['body'] = body
                    break
                elif part.get_content_type() == 'text/html':
                    body = part.get_payload(decode=True).decode(part.get_content_charset())
                    item['body'] = re.sub("<.*?>", "", body)
                    item['html'] = body
                    break
        else:
            body = email_message.get_payload(decode=True).decode(email_message.get_content_charset())
            item['body'] = body
        return item  # 返回邮件信息

    # 获取所有邮件
    def fetch_all_emails(self, search_type='ALL'):
        self.search_mail(search_type=search_type)

    # 根据发件人搜索邮件
    def fetch_emails_from_sender(self, text):
        self.search_mail(search_type=f'(FROM "{text}")')

    # 根据日期搜索邮件
    def fetch_emails_since_date(self, dt):
        text = datetime.strptime(dt, "%Y-%m-%d").strftime("%-d-%b-%Y")
        self.search_mail(search_type=f'(SINCE "{text}")')

    # 根据主题搜索邮件
    def fetch_emails_by_subject(self, text):
        datas = self.search_mail(search_type=f'(SUBJECT "{text}")')
        for data in datas:
            return data

    # 根据邮件正文搜索邮件
    def fetch_emails_by_body(self, text):
        datas = self.search_mail(search_type=f'(BODY "{text}")')
        for data in datas:
            return data

    # 执行邮件搜索操作
    def search_mail(self, search_type):
        for mailbox in ["Junk", 'inbox']:
            self.mail.select(mailbox)

            status, email_ids = self.mail.search(None, search_type)  # 执行搜索
            if status != 'OK':
                print("Error search email")
                exit()
            email_id_list = self.get_email_id_list(email_ids)[::-1]  # 获取邮件ID列表
            # 打印每封邮件的详细信息
            for email_id in email_id_list:
                email_info = self.get_email_info(email_id)
                yield email_info
