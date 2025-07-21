import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from contentMaking import *
import time

def send_emails(account,data,Recipients,settings = {},threadID = '主',
                stop_event=None,workbook=None,key_column=None,sheet_lock=None):
    '''
    建立SMTP并发送邮件的核心代码
    Args:
        account:用户设定的账户名
        data:字典,直接读取配置文件,json格式如下:
            {"sender":"bilicoco29@gmail.com",
            "senderName":"coco",
            "password":"abcdefghijklmnop",
            "SMTP_SERVER":"smtp.gmail.com",
            "SMTP_PORT":465}
            
        Recipients:收件人信息字典
            {'tonyp8': {'mail': 'tonyp8@126.com', 'region': 'DE','row':1},
             'lico': {'mail': 'bc29@gmail.com', 'region': 'DE','row':2}}
        
        message:字典,邮件内容,格式如下: 弃用
            {'subject':'标题',
            'content':'正文',
            'attachment':['附件目录']}
        settings:字典，config.json[settings]
        threadID:线程ID
        stop_event:终止线程信号
        workbook:表格
        key_column:发送状态列 F:15
        sheet_lock:进程锁
        
    '''
    #print(account,data,Recipients)
    threadID = threadID+'线程: '
    print(f'[INFO] {threadID}即将开始发送,等待3s')
    time.sleep(3)
    if stop_event and stop_event.is_set():
        print(f'[WARNING] 用户终止了线程:{threadID}')
        return    
    try:
        waitingTime = settings["intervalSendingTime"]
        if waitingTime <= 60:#防止系统超时登出
            # 创建 SMTP 连接
            server = smtplib.SMTP_SSL(data['SMTP_SERVER'], data['SMTP_PORT'])
            print(f'[DEBUG] {threadID}连接创建成功')
            server.login(data['sender'], data['password'])
            print(f'''[INFO] {threadID}{account}({data['sender']}) 登录成功''')
        
        #print(Recipients)
        for recipient in list(Recipients.keys()): #收件人
            if stop_event and stop_event.is_set():
                print(f'[WARNING] 用户终止了线程:{threadID}')
                return
            
            if waitingTime > 60:#防止系统超时登出
                # 创建 SMTP 连接
                print(f'[INFO] {threadID}正在登录')
                print(f'[DEBUG] {threadID}少女祈祷中...')
                server = smtplib.SMTP_SSL(data['SMTP_SERVER'], data['SMTP_PORT'])
                print(f'[DEBUG] {threadID}连接创建成功')
                server.login(data['sender'], data['password'])
                print(f'''[INFO] {threadID}{account}({data['sender']}) 登录成功''')
            
            # 构建邮件内容
            msg = MIMEMultipart()
            msg['From'] = data['sender']
            msg['To'] = Recipients[recipient]['mail']
            
            senderNameSettings = settings['senderName']
            if senderNameSettings['statue'] == 0: #发送人署名设置设置选项:
                #{0:'替换名字',1:'按照配置文件署名',2:'按照表格署名'}
                senderName = senderNameSettings['senderName']
            elif senderNameSettings['statue'] == 1:
                senderName = data['senderName']
            else:
                print(f'[WARNING] {threadID}发送人署名设置错误！')
                senderName = 'WESLAMIC Team'

            print(f'[DEBUG] {threadID}构建邮件内容')
            msg['Subject'],content,attachment = contentMaking(recipient,Recipients[recipient]['region'],senderName,settings['mailModelContent'])

            #print('[DEBUG] 邮件原始内容:',msg['Subject'],content,attachment)
            #添加正文
            msg.attach(MIMEText(content, 'plain', 'utf-8'))

            # 添加附件
            for file_path in attachment:
                with open(file_path, 'rb') as f:
                    file_part = MIMEApplication(f.read())
                    file_part.add_header('Content-Disposition', 'attachment', filename=file_path.split('/')[-1])
                    msg.attach(file_part)
            print(f'[DEBUG] {threadID}邮件内容构建完成')
            # 发送邮件
            #server.sendmail(data['sender'], Recipients[recipient]['mail'].replace(" ", ""), msg.as_string())
            print(f"[INFO] {threadID}邮件已成功发送至 {recipient}")
            
            if workbook and key_column:
                with sheet_lock:  # 获取锁
                    print(f"[DEBUG] 更新表格信息")
                    workbook.active.cell(row=Recipients[recipient]['row'], column=key_column).value = f"{senderName}已发送"
                    workbook.save(settings['content'])
                    print(f"[DEBUG] 更新表格信息成功")

            #优化终止线程判断
            ttime = time.time()
            while time.time() < ttime + waitingTime:
                if stop_event and stop_event.is_set():
                    print(f'[WARNING] 用户终止了线程:{threadID}')
                    return  
                time.sleep(1)
            
            print(f'[INFO] {threadID}等待{waitingTime}s\n')
            if waitingTime > 60:#防止系统超时登出
                server.quit()
            
        if waitingTime <= 60:
            server.quit()
            
        print(f'''[INFO] {threadID}{account}({data['sender']}) 登出成功''')
    except Exception as e:
        if str(e) == '[Errno 0] Error':
            print('[ERROR] 请检查网络连接')
        print(f"[ERROR] 邮件发送失败: {e}")

if __name__ == "__main__":
    print('[WARNING] 你正在运行发送核心！')
    print('[WARNING] 本核心单独运行报错！')
    # 配置发件人信息
    SMTP_SERVER = 'smtp.gmail.com'
    SMTP_PORT = 465
    SenderEmail = 'WESLAMIC.DE.Official@gmail.com'
    Password = 'abcdefghijklmnop"'

    # 收件人列表
    Recipients = ['tonyp8@126.com']

    # 邮件正文和附件路径
    Subject = 'Python 批量发送邮件测试'
    Content = '这是使用 Python 自动化发送的测试邮件。'
    #Attachment = ['path/to/attachment1.txt', 'path/to/attachment2.pdf']
    Attachment = []


    data = {'sender': SenderEmail, 'password': Password,
            'SMTP_SERVER': SMTP_SERVER, 'SMTP_PORT': SMTP_PORT}
    message = {'subject':Subject,'content':Content,'attachment':Attachment}

    send_emails(data,Recipients)
