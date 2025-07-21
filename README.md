# batch-Send-Email 批量发送邮件 操作指南（new）
batch Send Email .一个批量发送邮件的自动化脚本.支持从表格中筛选发送人，提取收信人名称、地区、地址，支持根据地区发送不同的邮件，支持自定义修改发件人、收件人姓名（通过替换邮件模板的特殊标记实现）

> _本程序是为了达人营销触达制作的，有些地方与达人营销有关，但是不影响使用_

源文件介绍：

- contentMaking.py       生成邮件body的代码, 可以添加附件(attachment), 在列表中输入文件目录即可. 但程序中**并没有**提供其传入参数_(因为我用不上)_
- sendMailCore.py         发送邮件的核心代码, 在先前的版本中可以单独运行,但现在不行
- sendMailMain.py         独立运行时是命令行窗口,通过修改配置文件实现功能. 能用,但不支持日志记录,在表格里记录成功发送信息等功能.是程序的精简版
- sendMailUI3.py           由DS辅助编写的UI,并且添加了如预览和修改模板, 及时终止线程, 实时编辑配置等等有用的功能.**推荐使用**
- setup.py                      没什么可说的,cx_Freeze的打包UI3的脚本. 不过实测下来pyinstaller体积更小一点,但代价是命令很长(需要连同下面两个文件一起打包)
- config.cfg                    原始配置文件, 不要编辑这个,编辑生成的config.json.
- DefaultSheet.sheet      默认的表格文件,打开文件会复制一份变成 expertSheet.xlsx. 可以通过配置文件修改成别的表格


首次设置时间大概在5分钟，开启STMP服务可能需要5-15分钟

初次运行请先打开sendMailUI3.py，点击“开始发送”，待提示“\[ERROR\] 发送过程中出错: local variable 'workbook' referenced before assignment”时即可关闭

<img width="1102" height="782" alt="Image" src="https://github.com/user-attachments/assets/4d737b7f-5915-46f4-9b14-03a8598ff905" />

> _初次打开并按下“开始发送”显示的界面_

## 1.配置config文件

### 1.程序设置

使用记事本等工具打开'config.json'

打开后的界面如下

<img width="1008" height="684" alt="Image" src="https://github.com/user-attachments/assets/2a87438e-0215-4e37-a4df-f93f48960025" />

需要手动修改的地方：

#### 1，邮件模板位置

<img width="533" height="179" alt="Image" src="https://github.com/user-attachments/assets/1d3b0ec3-09e3-4a3f-95d4-afa6896a5c2c" />

其中 "./email"表示程序运行目录下“email”文件夹，‘./email/EN.txt’即该文件夹中EN.txt文件。邮件模板的格式见下文

此项与达人地区关联，例如达人地区为“EN”，则使用‘./email/EN.txt’作为模板；若达人地区为IT，但模板位置并未填写该地区，则默认按照第一个（‘EN’）的模板发送。

**配置该项时需要注意不要漏了逗号**，按图片举例，仅有最后一行（‘DE’）末尾无需逗号，而上方其他项末尾都要加**英文**逗号

#### 2.筛选表格负责人设置

<img width="590" height="244" alt="Image" src="https://github.com/user-attachments/assets/7ca99e7f-e68e-4d7d-b951-67bcd924ecf8" />

此项功能用于筛选达人。例如如上图情况，选择0，完全匹配，“chargerName”填写coco，则仅筛选表格中'COCO'负责的达人来发送邮件。“忽略大小写”指的是“chargerName”中无论填 coco ，COCO ，CoCo ……都可以筛选出表格中 “COCO” 负责的达人. 需要注意的是,本功能要求表格中负责人名字是大写的, 源码(Main)中只写了一层upper.

#### 3.发送人署名设置

<img width="767" height="187" alt="Image" src="https://github.com/user-attachments/assets/a8ea5b98-0bbe-4707-95b3-53a5048c02d3" />

statue填 0 表示将邮件署名强制替换为 senderName ；填1时邮件署名与账号署名相同，详情见下文账号部分

### 2.账号设置

<img width="649" height="396" alt="Image" src="https://github.com/user-attachments/assets/13d10506-120a-4ded-974f-d597502186d7" />

按照注释**在下方修改**修改即可，需要注意的是，账户显示名称与信件署名**不同**，账户显示名称只是为了自己区分，对方看不到；信件署名对方可以看到。账户显示名称**不可以**相同，信件署名可以相同。

设置账户时注意，请严格按照json编辑配置文件，注意增删逗号！

#### 2.1获取SMTP服务密码

可以参考这篇文章：

[https://blog.csdn.net/weixin_54689482/article/details/146033946](url)

文章比较详细，按照步骤走就可以了

#### 2.2多账户设置

> 如果你有两个及以上的邮箱可以发送邮件，可以看这一节

> 你也可以打开软件进行配置，不需要手动改配置文件，但仍建议你阅读。

回到软件设置部分

<img width="649" height="268" alt="Image" src="https://github.com/user-attachments/assets/914fd105-01cf-436a-861a-a105797b3973" />

**各项功能说明**：

##### 2.2.1.交错发送

配置时填0表示**不关闭**交错发送，1表示**关闭**交错发送

该功能指的是用多个邮箱发送邮件时，多个账号互相独立发送邮件，程序运行顺序如下图：

<img width="589" height="482" alt="Image" src="https://github.com/user-attachments/assets/59c1fbd5-137c-409f-9162-af5416301e30" />

不打开则是：

<img width="534" height="505" alt="Image" src="https://github.com/user-attachments/assets/bcbb99ca-c117-4b4b-aae5-366e61471cf4" />

##### 2.2.2.单邮箱发送设置

如果你已经填写了两个以上的邮箱，但是想只启用一个，可以按这一节配置

<img width="510" height="110" alt="Image" src="https://github.com/user-attachments/assets/bbafa7a4-dbdb-4814-8973-1f42b227334f" />

启用该设置：填0表示禁用，1表示启用

启用后账户名：填写你希望使用的**账户显示名称，**不填或错填程序默认使用配置中第一个账户

## 2.邮件模板

邮件请使用下列格式：

<img width="168" height="164" alt="Image" src="https://github.com/user-attachments/assets/fb7d047e-21e1-4fb4-955a-d9da314434a9" />

标黄的第一行为邮件**标题**，留空则没有标题

第二行及下文则是邮件正文，

\[recName\] 指收件人名称（Recipient Name）

\[senderName\] 为发送者署名

以上两个均可不填写。

你也可以使用模板预览功能来辅助修改

## 3.填写发送达人

打开程序文件夹下的‘expertSheet.xlsx’（若自己设置了目录则按目录来）

复制达人建联表，并按照提示粘贴

其中**TKID必填，**负责人、达人地区、email可视情况不填但**数据必须准确**。若达人没有email的可以留空，但是**不允许**一格中同时填写邮箱和其他联系方式（程序仅通过“@”来判断是否存在邮件地址）。

O列（第15列）’邮件发送情况‘，留空，无需填写。该列是给程序判断是否已经发送了该邮箱。若填写则**不会发送**邮件给该达人。若邮件发送成功，程序会自动填写“发送账户署名+‘已发送’”。

## FinallStep

双击运行sendMailUI3.exe，即可开始批量发送邮件。初次发送建议先发给自己。

> 可能仍有些地方不够详实，向我提出我马上补充
