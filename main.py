import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.header import Header
import tkinter as tk
from tkinter import filedialog

# 初始化窗口
root = tk.Tk()
root.title("学生成绩通知系统")
root.geometry("400x300")

sender = '1183298554@qq.com'  # 邮件发送账号
reciever = '332582281@qq.com'  # 接收邮件账号
password = 'gqraeiwuiommbahc'  # 授权码（这个要填自己获取到的）
smtp_server = 'smtp.qq.com'  # 固定写死
smtp_port = 465  # 固定端口

#配置服务器
stmp=smtplib.SMTP_SSL(smtp_server,smtp_port)
stmp.login(sender,password)

def import_excel():#导入Excel文件
    global filename
    filename = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx *.xls')])
    if filename:
        print("已选择文件：", filename)
        df = pd.read_excel(filename, skiprows=[0])
        print(df)
    else:
        print("请选择一个Excel文件")

def send_email():
    try:
        df = pd.read_excel(filename, skiprows=[0])
    except FileNotFoundError:
        print("文件不存在或路径错误")
        exit()
    except Exception as e:
        print("发生了错误:", str(e))
        exit()

    grouped = df.groupby("姓名")
    contents = []
    for name, group in grouped:
        content = "亲爱的" + name + "同学：" + "\n"
        content += ("祝贺您顺利完成本学期的学习！教务处在此向您发送最新的成绩单。\n\n")
        for index, row in group.iterrows():
            content += ("[" + str(row["课程名称"]) + "]" + ":" + "[" + str(row["百分成绩"]) + "]" + "[" + str(row["五分成绩"]) + "]"+ "[" + str(row["考试类型"]) + "]"+ "[" + str(row["选修类型"]) + "]"+ "\n")

        content += (
            "希望您能够对自己的成绩感到满意，并继续保持努力和积极的学习态度。如果您在某些科目上没有达到预期的成绩，不要灰心，这也是学习过程中的一部分。我们鼓励您与您的任课教师或辅导员进行交流，他们将很乐意为您解答任何疑问并提供帮助。请记住，学习是一个持续不断的过程，我们相信您有能力克服困难并取得更大的进步。\n\n")
        content += ("再次恭喜您，祝您学习进步、事业成功！\n\n")
        content += ("教务处\n")

        reciever = receiver_entry.get()  # 获取输入框中的邮箱地址

        
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

def exit_app():
    root.destroy()  # 关闭GUI窗口

# 创建输入框和标签
receiver_label = tk.Label(root, text="接受者邮箱：")
receiver_label.pack()
receiver_entry = tk.Entry(root)
receiver_entry.pack()

# 创建导入按钮和发送按钮
import_button = tk.Button(root, text="导入Excel文件", command=import_excel)
send_button = tk.Button(root, text="发送邮件", command=send_email)
exit_button = tk.Button(root, text="退出", command=exit_app)
import_button.pack()
send_button.pack()
exit_button.pack()

root.mainloop()  # 启动GUI主循环
