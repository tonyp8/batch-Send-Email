def contentMaking(recName,recRegion,senderName,mailContent):
    '''
    生成邮件主体内容
    Args:
        recName:接收人名称
        #recINFO:字典，接收人信息 F:{'mail': 'tonyp8@126.com', 'region': 'DE'}
        recRegion:接收人地区区号 F:'EN'
        senderName:发送人名称
        mailContent:字典，config.json[settings]['mailModelContent']
                {"//":"按语言筛选，若没有则按第一个发送",
			"EN":"./email/EN.txt","DE":"./email/DE.txt"},
        
    Return:
        subject,content,attachment 标题，正文，附件目录列表
    '''
    #print(recName,recINFO,senderName,settings,end="\n")
    #recRegion = recINFO['region']
    #print('[DEBUG] 分区邮件目录',mailContent)
    #del mailContent['//']
    mailRegionList = list(mailContent.keys())
    #print(mailRegionList)
    #选择邮件模板
    if recRegion in mailRegionList:
        mailRegion = recRegion
    else:
        mailRegion = mailRegionList[0] #取第一项为默认值

    #读取邮件模板
    try:
        with open(mailContent[mailRegion], "r", encoding="utf-8") as f:
            subject = f.readline().strip()
            content = replaceName(f.read(),recName,senderName)
    except Exception as e:
        print(f"[ERROR] 读取模板出错! 无法读取文件：{e}")
        print('[WARNING] 发送自动终止!')
    #subject=f'''你好{recName},来自{recRegion}的网友'''
    #content=f'''你的地区为{recRegion}\n{senderName}'''
    attachment=[]
    #print(subject,content,attachment)
    return subject,content,attachment
    
def replaceName(content,recName,senderName):
    '''邮件标题正文,发送人署名,接收人姓名'''
    #print(f'邮件正文{content}')
    if not "[recName]" in content:
        print('[WARNING] 未在模板中找到收件人姓名替换处!')
        print('[WARNING] 格式: [recName]')
    if not "[senderName]" in content:
        print('[WARNING] 未在模板中找到发送者署名替换处!')
        print('[WARNING] 格式: [senderName]')
    content = content.replace("[recName]",str(recName))
    content = content.replace("[senderName]",str(senderName))
    return content
    
if __name__ == "__main__":
    pass
