# ============================================================
# 文件名称: RsoftMail.py
# 模块功能: 实现仿真完成后通过邮箱发送即时通知，辅助自动流程管理
# 使用场景: 仿真耗时较长，通知用户及时获取仿真结果
# 使用方式:
#   RsoftMail('your_email@example.com')
# 技术说明:
#   - 使用 QQ 邮箱 SSL SMTP 协议发送邮件
#   - 发件人邮箱需开启“POP3/SMTP服务”并生成授权码
# 作者单位: 中南大学机电工程学院 郑煜教授课题组
# 撰写作者: 2022届硕士研究生 万强伟
# 联系方式: 1025459384@qq.com
# ============================================================

import smtplib
from email.mime.text import MIMEText
from email.utils import formataddr

class RsoftMail:
    # === 构造函数: 初始化邮件参数并尝试发送通知邮件 ===
    # 函数名: __init__
    # 功能:
    #   - 创建邮件内容
    #   - 登录 SMTP 服务器
    #   - 发送至指定接收邮箱
    # 参数:
    #   yourMail - 接收者邮箱地址（字符串）
    # 返回:
    #   控制台输出发送结果（成功/失败）
    def __init__(self, yourMail=str):
        # 发件人邮箱（建议换为你自己的）
        self.my_sender = '2571277215@qq.com'

        # 发件人邮箱授权码（非登录密码，在邮箱设置中开启 SMTP 后生成）
        self.my_pass = 'nmiyoctyiappdjbc'

        # 接收人邮箱地址
        self.yourMail = yourMail

        send = True  # 标志变量，记录邮件是否成功发送

        try:
            # === 创建邮件内容 ===
            msg = MIMEText('Rsoft仿真完毕,你可以进行下一步处理', 'plain', 'utf-8')
            msg['From'] = formataddr(["RsoftMail", self.my_sender])  # 设置发件人昵称与地址
            msg['To'] = formataddr(["MyFriend", self.yourMail])      # 设置收件人昵称与地址
            msg['Subject'] = "RsoftMail"                         # 邮件主题

            # === 配置 SMTP_SSL 服务器 ===
            server = smtplib.SMTP_SSL("smtp.qq.com", 465)  # QQ邮箱SMTP服务器，SSL端口465
            server.login(self.my_sender, self.my_pass)     # 登录 SMTP 服务器
            server.sendmail(self.my_sender, [self.yourMail], msg.as_string())  # 发送邮件
            server.quit()                                  # 关闭连接

        except Exception:
            # 捕获异常（如网络、认证等错误），设置失败标志
            send = False

        # === 输出结果 ===
        if send:
            print("邮件发送成功")
        else:
            print("邮件发送失败")
