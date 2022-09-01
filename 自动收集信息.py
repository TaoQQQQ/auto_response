from wxauto import *
import schedule
import time
import datetime
import pandas as pd


data_init = pd.read_excel('C:/Users/Administrator/Desktop/test.xlsx',sheet_name='Sheet1',header=None)
data_one = pd.read_excel('C:/Users/Administrator/Desktop/test.xlsx',sheet_name='Sheet1',header=None)[1:]
data_two = pd.read_excel('C:/Users/Administrator/Desktop/test.xlsx',sheet_name='Sheet2',header=None)[1:]
init_time = str(datetime.datetime.now())[11:16]
print(init_time)
global last_time,lenth,already_write,last_wirte
last_time = init_time
last_wirte = []
already_write = data_init.shape[0]
lenth = data_two.shape[0]


# 比一下时间的大小
# 如果现在读取的时间比上一次大，就接收信息
def compare(time_now,last_time):
    #print('compare:time_now',time_now,'last_time',last_time)
    if len(time_now) > len(last_time):
        return 1
    else:
        time_now = eval(time_now[:time_now.find(':')])*60 + eval(time_now[time_now.find(':')+1:])
        last_time = eval(last_time[:last_time.find(':')])*60 + eval(last_time[last_time.find(':')+1:])
       # print('time_now',time_now,'last_time',last_time)
        if time_now >= last_time:
            return 1
        else:
            return 0

def get_information():
   data = {}
   init_data = list(data_one.iloc[:,1])
   for value in init_data:
       data[value] = data.get(value,0) + 1
   data_len = list(data.values())
   new_already_write = already_write - 1
   new_num = 0
   return_data = {}
   return_len = []
   member = {}
   for i in range(len(data_len)):
       old_num = new_num
       new_num = old_num + list(data.values())[i]

       flag = 0
       count = 0
       for j in range(old_num ,new_num):
           if data_one.iloc[j,9] != "是":
               flag = 1
               count = count + 1
               if data_one.iloc[j,5] not in list(member.keys()):
                 member[data_one.iloc[j,5]] = [data_one.iloc[j,0]]
               else:
                 member[data_one.iloc[j,5]].append(data_one.iloc[j,0])
       if flag == 1:
           return_data[list(data.keys())[i]] = list(data.values())[i]
           return_len.append(count)
           count = 0

   return return_data,return_len,member

def write_information(number):
    global already_write
    #print(data_one.loc[34,0])
    data_one.loc[already_write,0] = str(number)
    time = str(datetime.datetime.now())[5:10]
    if time[0] == '0':
        time = time[1:]
    #print('write_information:',time)
    data_one.loc[already_write, 1] = time
    for i in range(lenth):
      if number == str(data_two.iloc[i,0]) :
        data_one.loc[already_write, 5] = data_two.iloc[i,1]
        data_one.loc[already_write, 6] = data_two.iloc[i,5]
        data_one.loc[already_write, 7] = data_two.iloc[i,4]
        break
    print('写入人员为：',number)

    tmp = pd.DataFrame.copy(data_one)
    for i in range(data_one.shape[0], 1, -1):
        tmp.loc[i + 1, :] = tmp.loc[i, :]
    tmp.loc[1,] = data_init.iloc[0,]

    tmp.to_excel('C:/Users/Administrator/Desktop/the_test.xlsx', header=False, index=False)
    already_write = data_one.shape[0] + 1
# 每隔一个小时发送群消息
def send_message():
    print("start send!")
    msg = ""
    datas,count,members = get_information()
    data = list(datas.keys())
    lenth = list(datas.values())
    name = list(members.keys())
    number = list(members.values())
    for day in range(len(data)):
        msg = msg + str(data[day]) + " 预警客户共计" + str(lenth[day]) + "个，已反馈" + str(lenth[day]-count[day]) + "个，" + "未反馈" + str(count[day]) + "个" +  ','
    msg = msg + "具体为" + '\n'
    for member in range(len(name)):
        msg = msg + str(name[member]) + "：" + "、".join('%s' %x for x in number[member]) + "、"
    msg = msg + " 请以上人员及时反馈"
    print(msg)
    wx.ChatWith(group_name)
    WxUtils.SetClipboard(msg)    # 将内容复制到剪贴板，类似于Ctrl + C
    wx.SendClipboard()   # 发送剪贴板的内容，类似于Ctrl + V

# 每隔十分钟接收一次消息
def update():
    wx.ChatWith(group_name)
    wx.LoadMoreMessage()
    msgs = wx.GetAllMessage
    flag = 0
    global data_one
    for msg in msgs:
        if msg[0] == 'Time' and len(msg[1]) <= 5 and compare(str(msg[1]), last_time):
            flag = 1
            continue
        if flag == 1:
            print('start update!')
            message = msg[1].split('\n')
            if len(msg[1]) >= 22 and message[0][:6] == "【用户姓名】" and message[1][:4] == "【号码】" and message[2][ :8] == "【是否挽留成功】" and message[3][:4] == "【原因】":

                phone_number = str(message[1][4:])
                succeed = str(message[2][8:])
                reason = str(message[3][4:])
                print("update_phonenumber:",phone_number)
                for j in range(already_write-1):
                  if str(data_one.iloc[j, 0]) == phone_number:
                    data_one.loc[j + 1, 9] = "是"
                    data_one.loc[j + 1, 10] = succeed
                    data_one.loc[j + 1, 11] = reason

    tmp = pd.DataFrame.copy(data_one)

    for i in range(data_one.shape[0], 1, -1):
        tmp.loc[i + 1, :] = tmp.loc[i, :]
    tmp.loc[1,] = data_init.iloc[0,]
    tmp.to_excel('C:/Users/Administrator/Desktop/the_test.xlsx', header=False, index=False)

def receive():
    print('start receive!')
    wx.ChatWith(mentor_name)
    msgs = wx.GetAllMessage
    flag = 0
    global last_time,already_write,last_wirte
    for msg in msgs:
        if msg[0] == 'Time' and len(msg[1]) <= 5 and compare(str(msg[1]),last_time):
            flag = 1
            last_time = msg[1]
            # print('receive: last_time:',last_time)
            continue
        if flag == 1:
            #  print('receive: ',msg[1])
            test = msg[1][21:32]
            if len(msg[1]) >= 8 and '【携转预警提醒】' == msg[1][:8] :
                # 把号码整下来就行
                phone_number = msg[1][msg[1].find('1'): msg[1].find('1') + 11]
                if phone_number not in last_wirte:
                    phone_number = msg[1][msg[1].find('1') : msg[1].find('1') + 11]
                    print("phone_number:",phone_number)
                    write_information(phone_number)
                    last_wirte.append(phone_number)
                    print('写入!')
    update()


wx = WeChat()
mentor_name = '文件传输助手'
group_name = '测试群'
wx.GetSessionList()
wx.ChatWith(mentor_name)
msgs = wx.GetLastMessage

schedule.every(0.2).minutes.do(receive)
schedule.every(0.2).minutes.do(send_message)


while True:
    schedule.run_pending()



