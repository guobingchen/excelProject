陌上花开 2023/10/22 15:50:36
那玩家属性里要加个房间号吗

飞翔 2023/10/22 15:50:58
不用吧

飞翔 2023/10/22 15:51:11
不太懂

你撤回了一条消息

你撤回了一条消息

陌上花开 2023/10/22 16:01:38


import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.header import Header
import os

sender='1183298554@qq.com'  #邮件发送账号
password='gqraeiwuiommbahc'  #授权码（这个要填自己获取到的）
smtp_server='smtp.qq.com'#固定写死
smtp_port=465#固定端口
 
 
#配置服务器
stmp=smtplib.SMTP_SSL(smtp_server,smtp_port)
stmp.login(sender,password)

current_path = os.path.dirname(os.path.abspath(__file__))
print(current_path)

filename = input("请输入Excel文件名：")

# 获取当前文件所在目录的绝对路径
current_path = os.path.dirname(os.path.abspath(__file__))

# 拼接Excel文件的完整路径
excel_path = os.path.join(current_path, 'Excel', filename)

try:
    df = pd.read_excel(filename, skiprows=[0])
    # 在这里进行后续的数据处理操作
except FileNotFoundError:
    print("文件未找到，请检查文件名和路径是否正确。")
    exit()
except Exception as e:
    print("发生了错误:", str(e))
    exit()

# 按照姓名分组
grouped = df.groupby("姓名")

# 创建一个空数组用于存储每个分组的内容
contents = []

# 遍历每个分组
for name, group in grouped:
    # 创建一个新的字符串用于存储当前分组的内容
    content = "亲爱的" + name +"同学：" + "\n"
    content +=("祝贺您顺利完成本学期的学习！教务处在此向您发送最新的成绩单。\n\n")
    
    # 遍历该组的每一行数据
    for index, row in group.iterrows():
        # 输出课程名称、学分、百分成绩、五分成绩
        content += ("[" + str(row["课程名称"]) + "]" + ":"
                    "[" + str(row["百分成绩"]) + "]" + "\n")
    content += ("希望您能够对自己的成绩感到满意，并继续保持努力和积极的学习态度。如果您在某些科目上没有达到预期的成绩，不要灰心，这也是学习过程中的一部分。我们鼓励您与您的任课教师或辅导员进行交流，他们将很乐意为您解答任何疑问并提供帮助。请记住，学习是一个持续不断的过程，我们相信您有能力克服困难并取得更大的进步。\n\n")
    
    content += ("再次恭喜您，祝您学习进步、事业成功！\n\n")

    content +=("教务处\n")

    reciever=input("请输入接受者的邮箱：")

    message = MIMEText(content, 'plain', 'utf-8')  #发送的内容
    message['From'] = sender
    message['To'] = reciever
    subject = '成绩单发送'
    message['Subject'] = Header(subject, 'utf-8') #邮件标题

    try:
        stmp.sendmail(sender, reciever, message.as_string())
    except Exception as e:
        print ('邮件发送失败--' + str(e))
        print ('邮件发送成功')


