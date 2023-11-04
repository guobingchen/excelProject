陌上花开 2023/10/22 15:50:36
那玩家属性里要加个房间号吗

飞翔 2023/10/22 15:50:58
不用吧

飞翔 2023/10/22 15:51:11
不太懂

你撤回了一条消息

你撤回了一条消息

陌上花开 2023/10/22 16:01:38


飞翔 2023/10/22 16:07:42
建筑和兵种都要有生命值，攻击，防御，攻击范围，攻击间隔，费用，因为建筑可以当成不会移动的兵种，等级就不要了，兵种在加个移动速度

陌上花开 2023/10/22 16:08:13
okk

陌上花开 2023/10/22 16:08:22
那其他还有什么要加的吗

飞翔 2023/10/22 16:09:39
等级都不要了，我们应该会写死，建筑也要有占地范围，兵种去掉“死亡掉落经验”

陌上花开 2023/10/22 16:14:34
防御要分成防御范围和能力吗

飞翔 2023/10/22 16:15:21
不用，就防御

你撤回了一条消息

陌上花开 2023/10/22 16:28:50


陌上花开 2023/10/22 16:29:02
这样可以吗

飞翔 2023/10/22 16:29:23
等级是全都不要了

飞翔 2023/10/22 16:29:45
我们只做演示的话，不会让玩家有时间升级的

陌上花开 2023/10/22 16:30:10
噢噢这样

飞翔 2023/10/22 16:30:23
其他没问题了

飞翔 2023/10/30 22:40:38


飞翔 2023/10/30 22:40:55
这里有个问题就是不是卡牌直接改变玩家分数

飞翔 2023/10/30 22:41:19
你这里的分数是玩家游戏内的生命值之类的概念吗

飞翔 2023/10/30 22:41:36
还是打出卡牌需要的费用

陌上花开 2023/10/30 22:41:39
不是 就是总体成绩

陌上花开 2023/10/30 22:41:50
那应该咋改

飞翔 2023/10/30 22:42:08
就排行榜之类的分数？

陌上花开 2023/10/30 22:42:26
对

飞翔撤回了一条消息

陌上花开 2023/10/30 22:43:19
就是玩家类里面应该有自己的分数吧 游戏结束之后再更新数据库 我是这样想的

飞翔 2023/10/30 22:44:15
现在这样像是直接用卡牌攻击防御，但是这是攻击法术类卡牌的功能，像是兵种和建筑都不算是卡牌直接攻击防御

飞翔 2023/10/30 22:44:46
8那个改一下，改成卡牌可用时打出卡牌吧

陌上花开 2023/10/30 22:44:51
okk

飞翔 16:42:57
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

该用户通过“软工”群向你发起临时会话，前往设置。

