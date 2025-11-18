# 主群处理
#.1先查没收到的，无赞藏sheet，看是不是正确，有的人发的小红书号不对，如少发了个特殊字符，这时候给他们加回去。
#2到sheet1里查有没有“关注”字眼，给他加上钱，因为不自动计算。
#3.查小红书的 “按小红书查到的计算”列和“按用户发的然后从小红书查找计算”列对比的结果列“按小红书查到的计算==按用户发的然后从小红书查找计算”结果为false的，看看为什么
#3.1如果不是true，看是否有不同人但是小红书号名字一样，导致给他多结算了，要按照用户发的结算
#4.看无支付码sheet，里是不是有人已经发支付码了
from wxauto4  import WeChat
from  wxauto4 import  uia
import xlwings as xw
import sqlite3 
import datetime
import os
import traceback
from PIL import Image
import time
import csv
import re
import ctypes
import threading
class DHMsg:
    def __init__(self, type, myType,sender,content,time,fromw):
        self.type = type
        self.myType = myType
        self.sender=sender
        self.content=content
        self.time=time
        self.fromw=fromw 
def xor_encrypt_decrypt(text, key=1123):
    encrypted_text = ""
    for char in text:
        encrypted_text += chr(ord(char) ^ key)
    return encrypted_text
def add_center(sht, target, filePath, match=False, width=None, height=None, column_width=None, row_height=None):
    '''Excel智能居中插入图片

    优先级：match > width & height > column_width & row_height
    建议使用column_width或row_height，定义单元格最大宽或高

    :param sht: 工作表
    :param target: 目标单元格，字符串，如'A1'
    :param filePath: 图片绝对路径
    :param width: 图片宽度
    :param height: 图片高度
    :param column_width: 单元格最大宽度，默认100像素，0 <= column_width <= 1557.285
    :param row_height: 单元格最大高度，默认75像素，0 <= row_height <= 409.5
    :param match: 绝对匹配原图宽高，最大宽度1557.285，最大高度409.5
    '''
    unit_width = 6.107  # Excel默认列宽与像素的比
    rng = sht.range(target)  # 目标单元格
    name = os.path.basename(filePath)  # 文件名
    _width, _height = Image.open(filePath).size  # 原图片宽高
    NOT_SET = True  # 未设置单元格宽高
    # match
    if match:  # 绝对匹配图像
        width, height = _width, _height
    else:  # 不绝对匹配图像
        # width & height
        if width or height:
            if not height:  # 指定了宽，等比计算高
                height = width / _width * _height
            if not width:  # 指定了高，等比计算宽
                width = height / _height * _width
        else:
            # column_width & row_height
            if column_width and row_height:  # 同时指定单元格最大宽高
                width = row_height / _height * _width  # 根据单元格最大高度假设宽
                height = column_width / _width * _height  # 根据单元格最大宽度假设高
                area_width = column_width * height  # 假设宽优先的面积
                area_height = row_height * width  # 假设高优先的面积
                if area_width > area_height:
                    width = column_width
                else:
                    height = row_height
            elif not column_width and not row_height:  # 均无指定单元格最大宽高
                column_width = 100
                row_height = 75
                rng.column_width = column_width / unit_width  # 更新当前宽度
                rng.row_height = row_height  # 更新当前高度
                NOT_SET = False
                width = row_height / _height * _width  # 根据单元格最大高度假设宽
                height = column_width / _width * _height  # 根据单元格最大宽度假设高
                area_width = column_width * height  # 假设宽优先的面积
                area_height = row_height * width  # 假设高优先的面积
                if area_width > area_height:
                    height = row_height
                else:
                    width = column_width
            else:
                width = row_height / _height * _width if row_height else column_width  # 仅设了单元格最大宽度
                height = column_width / _width * _height if column_width else row_height  # 仅设了单元格最大高度
    assert 0 <= width / unit_width <= 255
    assert 0 <= height <= 409.5
    if NOT_SET:
        rng.column_width = width / unit_width  # 更新当前宽度
        rng.row_height = height  # 更新当前高度
    left = rng.left + (rng.width - width) / 2  # 居中
    top = rng.top + (rng.height - height) / 2
    try:
        sht.pictures.add(filePath, left=left, top=top, width=width, height=height, scale=None, name=name+str(len(sht.pictures)))
    except Exception:  # 已有同名图片，采用默认命名
        pass
def InsertMarkID(name,odata,MarkID):
#names:要插入的所有的名字。odata:已经在数据库的数据
        id=-1
        for od in odata:
            if(od[4]==name):
                id=od[0]
                break
        if(id==-1):
            print(f"**增加了新用户{name}，他MarkID是{MarkID+1}")
            insert_single_sql = '''INSERT INTO MarkWX (ID,wxName ,MarkID ,PayCode,OriginName,AddTime)
            VALUES (?,?, ?,?,?,?)'''  
            cursorsql.execute(insert_single_sql, (MarkID+1,name,MarkID+1,0,name,datetime.datetime.now().strftime("%Y_%m_%d_%H_%M_%S")))
            # 提交事务，将更改保存到数据库
            conn.commit() 
            return  MarkID+1
def CreateTableWxInfo():
    global cursorsql
    create_table_sql = '''
    CREATE TABLE IF NOT EXISTS WXHandleInfo (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        wxID TEXT NOT NULL,
        xhsID TEXT NOT NULL,
        IsZ INTEGER,
        IsC INTEGER,
        IsP INTEGER,
        ZhengMing TEXT,
        IsConfirm INTEGER,
        IsPay INTEGER 
    )
    '''
    # 执行 SQL 语句创建表
    cursorsql.execute(create_table_sql)
    # 提交事务，将更改保存到数据库
    conn.commit() 
def CreateTableWxToXHS():
    global cursorsql
    create_table_sql = '''
    CREATE TABLE IF NOT EXISTS WXToXHSInfo (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        wxID TEXT NOT NULL,
        xhsID TEXT NOT NULL ,
        AddTime TEXT NOT NULL 
    )
    '''
    # 执行 SQL 语句创建表
    cursorsql.execute(create_table_sql)
    # 提交事务，将更改保存到数据库
    conn.commit() 
def InsertWXInfoTocache(infos):
    #infos=[("大海"，“xiaoxiao”，1,1,0，“dsdfsd.mp4”,1,1)]
    #
    #
    #
    global cursorsql
    # 定义插入单条数据的 SQL 语句
    insert_single_sql = '''INSERT INTO WXHandleInfo (wxID ,xhsID ,IsZ , IsC, IsP, ZhengMing , IsConfirm ,IsPay,addTime,msg,contentAll)
     VALUES (?, ?,?,?, ?,?,?, ?,?, ?,?)'''
    # 插入单条数据
    cursorsql.executemany(insert_single_sql, infos)

    # # 定义插入多条数据的 SQL 语句
    # insert_multiple_sql = "INSERT INTO users (name, age) VALUES (?, ?)"
    # # 要插入的多条数据
    # data = [
    #     ('Bob', 30),
    #     ('Charlie', 35)
    # ]
    # # 批量插入数据
    # cursor.executemany(insert_multiple_sql, data)

    # 提交事务，将更改保存到数据库
    conn.commit()
def InsertWXToXHScache(infos):
    select_sql = "SELECT id,wxID,xhsID,AddTime,PayCode,MarkID FROM WXToXHSInfo"
    # 执行查询语句
    global cursorsql
    cursorsql.execute(select_sql)
    # 获取所有查询结果
    dataNode = cursorsql.fetchall() 
    toinsertInfo=[]
    for info in infos:
        find=False
        for dn in dataNode:
            if(dn[1]==info[0] and (dn[2] in info[1] or info[1] in dn[2])) or info[1]=="":
                find=True
                break
        if(find==False):
            findInToInsert=False
            for toinsert in toinsertInfo:
                if(toinsert[1]==info[0] and (toinsert[2] in info[1] or info[1] in toinsert[2])):
                    findInToInsert=True
                    break
            if(findInToInsert==False):
                toinsertInfo.append(info)
                #print(f"Wx-XHS增加了{info[0]}他备注为{info[3]}")

     
    # 定义插入单条数据的 SQL 语句datetime.datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
    insert_single_sql = '''INSERT INTO WXToXHSInfo (wxID ,xhsID ,AddTime,MarkID,PayCode)
     VALUES (?, ?,?,?,?)'''      
    cursorsql.executemany(insert_single_sql, toinsertInfo)
    # # 定义插入多条数据的 SQL 语句
    # insert_multiple_sql = "INSERT INTO users (name, age) VALUES (?, ?)"
    # # 要插入的多条数据
    # data = [
    #     ('Bob', 30),
    #     ('Charlie', 35)
    # ]
    # # 批量插入数据
    # cursor.executemany(insert_multiple_sql, data)

    # 提交事务，将更改保存到数据库
    conn.commit()
def InsertXMLNotReceive(infos):
    global wb ,sht3
    sht3.range('A1') .value=['备注号' ,'微信用户名','小红书名',"类型","备注","内容"] 
    i=2
    sht3.range('F:F').column_width = 50
    sht3.range('B:B').column_width = 20
    sht3.range('C:C').column_width = 20
    for info in infos: 
            sht3.range(f'A{i}') .value=list(info) 
            i+=1
    
    # wb.save(f'Result\\结算{datetime.datetime.now().strftime("%Y_%m_%d_%H_%M_%S")}.xls')
    # wb.close()
def InsertXML(infos):
    global sht,wb ,sht1,DZDay
    wxdata=[{}]#[{"wx":"wx用户名"，"xhs":"xhs用户名","z":"1","C":"1","P":"0","sp":"视频证明地址"}]
    # 将a1,a2,a3输入第一列，b1,b2,b3输入第二列
    header=['微信用户名','备注号' ,"按用户发的计算","按小红书查到的计算","按用户发的然后从小红书查找计算","按小红书查到的计算==按用户发的然后从小红书查找计算","支付码图","金额计算过程","操作账号个数","实际操作数","小红书账号","信息内容"] 
    sht1.range('A1') .value=header#没有支付码的人
    # sht1.range('F:F').column_width = 15
    # sht1.range('G:G').column_width = 15
    sht1.range('H:H').column_width = 40
    sht1.range('G:G').column_width = 40
    sht1.range('J:J').column_width = 40
    sht1.range('K:K').column_width = 15
    sht1.range('L:L').column_width = 40

    sht.range('A1') .value=header
    sht.range('A:A').column_width = 20
    sht.range('D:D').column_width = 20 
    sht.range('E:E').column_width = 20 
    sht.range('H:H').column_width = 40
    sht.range('J:J').column_width = 40
    sht.range('K:K').column_width = 15
    sht.range('L:L').column_width = 40 
    i=2
    sht1i=2
    nopaycode=""
    paycodeMarkID=""
    for info in infos: 
        path=f'config\\zfcode\\{int(info[3])}.jpg'
        insertList=list((info[0],info[1],info[2],info[7],info[8],None,info[3],info[4],info[5],info[6],info[9],info[10]))
        
        if(int(info[3])==0):#没有支付码的
            nopaycode+=f"@{info[0]}"
            paycodeMarkID+=f",{info[1]}"
            sht1.range(f'D{sht1i}').api.WrapText = True
            sht1.range(f'A{sht1i}') .value=insertList
            sht1.range(f'F{sht1i}').formula = f'=D{sht1i}=E{sht1i}'
            sht1i+=1

        else: 
            if(os.path.exists(path)):
                sht.range(f'D{i}').api.WrapText = True
                sht.range(f'A{i}') .value=insertList
                sht.range(f'F{i}').formula = f'=D{i}=E{i}'
                filePath = os.path.join(os.getcwd(),path )
                add_center(sht, 'G'+str(i), filePath, width=350, height=350)
                i+=1  
            else:
                print(f"{info[0]},MarkID{info[1]}在数据库有paycode但是没有文件")
                sht1.range(f'D{sht1i}').api.WrapText = True
                sht1.range(f'A{sht1i}') .value=insertList
                sht1.range(f'F{sht1i}').formula = f'=D{sht1i}=E{sht1i}'
                nopaycode+=f"@{info[0]}"
                paycodeMarkID+=f",{info[1]}"
                sht1i+=1
    if(nopaycode!=""):
        print(nopaycode)
        print(paycodeMarkID)
    print("未提供支付码的人：")
    wb.save(f'Result\\结算({min(DZDay).strftime("%d")}-{max(DZDay).strftime("%d")}){datetime.datetime.now().strftime("%Y_%m_%d_%H_%M_%S")}.xls')
    wb.close()
def LoadFromZFMulty():
    findindex=1
    msgs = []
    msgszfzf=[]#转发中的转发
    handledWnd=[]
    while(check_window_exists('ChatRecordWnd',findindex)):
        myconf=uia.WindowControl(ClassName='ChatRecordWnd', searchDepth=1,foundIndex=findindex)
        if(myconf.NativeWindowHandle not in handledWnd): 
            handledWnd.append(myconf.NativeWindowHandle)
        else:
            continue
        myconf.SwitchToThisWindow()
        ddsd2=myconf.GetChildren() 
        ds2=ddsd2[1].ListControl() 
        infosToSave=[]
        canload=True
        minus=9999
        while(canload):
            msgitems2=ds2.GetChildren()
            for MsgItem2 in msgitems2:
                if MsgItem2.ControlTypeName == 'ListItemControl' :
                        textbox1=MsgItem2.TextControl(searchDepth=10,foundIndex=1)
                        textbox2=MsgItem2.TextControl(searchDepth=10,foundIndex=2)
                        if(MsgItem2.Name==""  ):
                            try:
                                try:
                                    text_control = MsgItem2.TextControl(searchDepth=30,foundIndex=4)
                                    if text_control.Exists() and text_control.Name=="视频":
                                        continue
                                except Exception as ex1:
                                    pass
                                find=False
                                for mg in msgszfzf:
                                    if(mg.content==text_control.Name and mg.sender==textbox1.Name and mg.time== textbox2.Name):
                                        find=True
                                        break
                                date_time = datetime.datetime.strptime(textbox2.Name, "%m-%d %H:%M:%S")
                                # if(datetime.datetime(1900,3,16,0,0,0)>date_time):
                                #     continue
                                if(find==False):
                                    text_controlT=text_control.Name.split(": ")
                                    tc=text_control.Name.replace(text_controlT[0],"").replace("[视频]","").replace(": ","")
                                    msgszfzf.append(DHMsg("text","text",textbox1.Name,tc,textbox2.Name,MsgItem2.Name))         
                            except Exception as ex:
                                print(f"非转发")
                        if(MsgItem2.Name!="" and "[图片]" not in MsgItem2.Name):
                            textbox3=MsgItem2.TextControl(searchDepth=10,foundIndex=3)
                            find=False
                            for mg in msgs:
                                if(mg.content==textbox3.Name and mg.sender==textbox1.Name and mg.time== textbox2.Name):
                                    find=True
                                    break
                            date_time = datetime.datetime.strptime(textbox2.Name, "%m-%d %H:%M:%S")
                            # if(datetime.datetime(1900,3,16,0,0,0)>date_time):
                            #     continue
                            if(find==False):
                                msgs.append(DHMsg("text","text",textbox1.Name,textbox3.Name,textbox2.Name,MsgItem2.Name))
            resl= LoadMoreMessage(ds2,minus)
            canload=resl[0]
            minus=resl[1]
            time.sleep(1) 
        findindex+=1
     
    for msg in msgszfzf:
        findin=[ms for ms in msgs if (ms.sender==msg.sender and ("2组" in ms.content or "两组" in ms.content) )]
        if(len(findin)>0):
            findin[0].content+= msg.content
        
    return msgs
def checkwindowExist(window):
    return window.Exists(8,3)
def check_window_exists(class_name,findindex=1):
    try:
        # 尝试查找具有指定 ClassName 的 WindowControl
        window = uia.WindowControl(ClassName=class_name, searchDepth=1,foundIndex=findindex)
 
        # 如果找到了控件，确保它是有效的（没有被销毁）
        findresult=call_with_timeout(checkwindowExist, args=(window,), timeout=5)
        if findresult:
            return True
    except Exception as e:
        print(f"查找控件时出现错误: {e}")
    return False
def LoadFromZF():
    msgs = []
    if(check_window_exists('ChatRecordWnd')==False):
        return msgs
    myconf=uia.WindowControl(ClassName='ChatRecordWnd', searchDepth=1,foundIndex=4)
    ddsd2=myconf.GetChildren()
    ds2=ddsd2[1].ListControl()
    infosToSave=[]
    canload=True
    while(canload):
        msgitems2=ds2.GetChildren()
        for MsgItem2 in msgitems2:
            if MsgItem2.ControlTypeName == 'ListItemControl' :
                    if(MsgItem2.Name!="" and "[图片]" not in MsgItem2.Name):
                        textbox1=MsgItem2.TextControl(searchDepth=10,foundIndex=1)
                        textbox2=MsgItem2.TextControl(searchDepth=10,foundIndex=2)
                        textbox3=MsgItem2.TextControl(searchDepth=10,foundIndex=3)
                        find=False
                        for mg in msgs:
                            if(mg.content==textbox3.Name and mg.sender==textbox1.Name):
                                find=True
                                break
                        date_time = datetime.datetime.strptime(textbox2.Name, "%m-%d %H:%M:%S")
                        # if(datetime.datetime(1900,3,16,0,0,0)>date_time):
                        #     continue
                        if(find==False):
                            msgs.append(DHMsg("text","text",textbox1.Name,textbox3.Name,textbox2.Name,MsgItem2.Name))
        canload= LoadMoreMessage(ds2)
        time.sleep(1) 
    #myconf.SendKeys('{Esc}')
    return msgs

def LoadMoreMessage( C_MsgList: uia.ListControl,minnus=9999):
        """加载当前聊天页面更多聊天信息
        
        Returns:
            bool: 是否成功加载更多聊天信息
        """
        loadmore = C_MsgList.GetLastChildControl()
        loadmore_bottom = loadmore.BoundingRectangle.bottom
        bottom = C_MsgList.BoundingRectangle.bottom
        while True:
            if loadmore.BoundingRectangle.bottom < bottom and minnus!=bottom-loadmore.BoundingRectangle.bottom:#or loadmore.Name == ''
                minnus=bottom-loadmore.BoundingRectangle.bottom
                isload = True
                break
            else:
                C_MsgList.WheelDown(wheelTimes=5, waitTime=0.1)
                if loadmore.BoundingRectangle.bottom == loadmore_bottom or minnus==bottom-loadmore.BoundingRectangle.bottom:
                    isload = False
                    break
                else:
                    loadmore_bottom = loadmore.BoundingRectangle.bottom
        C_MsgList.WheelDown(wheelTimes=2, waitTime=0.1)
        return  (isload,minnus)
def remove_chars_around_colon(s):
    colon_index = s.find(':')
    if colon_index == -1:
        colon_index = s.find('：')
    if colon_index == -1:
        colon_index = s.find('.')
    if colon_index == -1:
        return s

    left_index = colon_index - 1
    left_count = 0
    # 向左查找最多两个数字
    while left_index >= 0 and left_count < 2 and s[left_index].isdigit():
        left_index -= 1
        left_count += 1

    right_index = colon_index + 1
    right_count = 0
    # 向右查找最多两个数字
    while right_index < len(s) and right_count < 2 and s[right_index].isdigit():
        right_index += 1
        right_count += 1

    # 拼接结果
    return s[:left_index + 1] + s[right_index:]
    # # 查找冒号的位置
    # colon_index = s.find(':')
    # if(colon_index<=0):
    #     colon_index = s.find('：')
    # # 如果没找到冒号，直接返回原字符串
    # if colon_index == -1:
    #     return s
    # # 检查冒号左右两边是否都至少有两个字符
    # if colon_index >= 2 and colon_index + 3 <= len(s):
    #     # 拼接去掉冒号及其左右各两个数字后的字符串
    #     return s[:colon_index - 2] + s[colon_index + 3:]
    # return s
def GetXHSID(xhsIDs,type):
    ##type:"赞""藏""评论"
    dataNodeDZ1FailedXHSID=[]
    dataNodeDZ1FailedXHSID= [data[3].replace(" ", "").replace("，","").replace(" ","").lower() for  data in xhsIDs if data[5]==type]#不符合要求的点赞数据，点上了，但不符合要求
    dataNodeDZ1FailedXHSID.extend([data.replace("小红薯","")  for data in  dataNodeDZ1FailedXHSID if "小红薯" in data])
    dataNodeDZ1FailedXHSID.extend([data.replace("用户","")  for data in  dataNodeDZ1FailedXHSID if "用户" in data])
    return dataNodeDZ1FailedXHSID

def parse_wechat_time(time_str):
    # 获取当前日期和时间
    now = datetime.datetime.now()
    today = now.date()
    year = now.year
    month = now.month

    # 正则表达式匹配不同格式
    patterns = [
        # 模式1：具体日期（如：2025年4月25日 5:48）
        r'^(\d{4})年(\d{1,2})月(\d{1,2})日 (\d{1,2}):(\d{2})$',
        # 模式2：星期+时间（如：星期一 9:40）
        r'^(星期一|星期二|星期三|星期四|星期五|星期六|星期天) (\d{1,2}):(\d{2})$',
        # 模式3：昨天+时间（如：昨天 9:40）
        r'^昨天 (\d{1,2}):(\d{2})$',
        # 模式4：仅时间（如：9:40，默认当天）
        r'^(\d{1,2}):(\d{2})$'
    ]

    for pattern in patterns:
        match = re.match(pattern, time_str)
        if match:
            if pattern == patterns[0]:
                # 模式1：解析年/月/日
                y, m, d, h, mi = match.groups()
                return datetime.datetime(int(y), int(m), int(d), int(h), int(mi))
            elif pattern == patterns[1]:
                # 模式2：解析星期（需计算对应日期）
                week_day = match.group(1)
                h, mi = match.group(2), match.group(3)
                # 计算当前星期几（0=星期一，1=星期二，...，6=星期日）
                current_weekday = now.weekday()  # 0=星期一，6=星期日
                target_weekday = ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期天'].index(week_day)
                # 计算距离今天的天数差（负数表示过去，正数表示未来）
                delta_days = target_weekday - current_weekday
                if delta_days < 0:
                    delta_days += 7  # 处理跨周情况
                target_date = today - datetime.timedelta(days=(7 - delta_days)) if delta_days != 0 else today
                return datetime.datetime(target_date.year, target_date.month, target_date.day, int(h), int(mi))
            elif pattern == patterns[2]:
                # 模式3：昨天的日期
                h, mi = match.groups()
                yesterday = today - datetime.timedelta(days=1)
                return datetime.datetime(yesterday.year, yesterday.month, yesterday.day, int(h), int(mi))
            elif pattern == patterns[3]:
                # 模式4：今天的时间
                h, mi = match.groups()
                return datetime.datetime(year, month, today.day, int(h), int(mi))
    # 若无法匹配，抛出异常或返回当前时间
    raise ValueError(f"无法解析的时间格式：{time_str}")
def loadMoreCleaver(AreaText):
    global StartText,breakText
    sText=AreaText[0]
    bText=AreaText[1]
    tempMsg=[] 
    toreturnMsg=[]
    needSetStartText= True if StartText==None else False
    needSetBreakText= True if breakText==None else False
    fastbreak=None#上一个上划中最小的时间，用于快速加载找到起始时间
    i=20
    while(i>0) :
        i-=1
        print(f"开始找起始时间{fastbreak}")
        tempMsg=wx.GetAllMessage(
                # savepic   = False,   # 保存图片
                # savefile  = False,   # 保存文件
                # savevoice = False,    # 保存语音转文字内容
                # saveVideo=False,
                # saveZF=False,
                # AreaText=(AreaText[0],fastbreak)
            ) 
        print("开始找起始时间加载结束")
        canbreak=False 
        fastbreak=None
        for msg in tempMsg:
            if msg.type == 'base'  :
                try:
                    parse_wechat_time(msg.content)
                except Exception as ex:
                    continue
                if sText > parse_wechat_time(msg.content).date() or (needSetStartText==False and StartText==msg.content):
                    canbreak=True
                else:
                    if(fastbreak==None):
                        fastbreak=msg.content
                    if(needSetStartText):
                        StartText=msg.content
                        break
        if (needSetBreakText==True):
            for msg in reversed(tempMsg):
                if msg.type == 'base'  :
                    try:
                        parse_wechat_time(msg.content)
                    except Exception as ex:
                        continue
                    if bText!= parse_wechat_time(msg.content).date():
                        if(needSetBreakText):
                            breakText=msg.content
                        continue
                    else:
                        needSetBreakText=False
                        print(f"*********找到结束时间了{breakText}")
                        break
        if(needSetBreakText==False):
            for msg in reversed(tempMsg):
                if msg.type == 'base'  :
                    try:
                        parse_wechat_time(msg.content)
                    except Exception as ex:
                        continue
                    if bText!= parse_wechat_time(msg.content).date():
                        pass
                    else:
                        break
                else:
                    toreturnMsg.append(msg)
        if(canbreak==True):
            print(f"*******找到起始时间了{StartText}") 
            break
        wx.LoadMoreMessage()
        # elif(wx.LoadMoreMessage()):
        #     print("向上滚动")   
        # else:
        #     print("加载信息失败，可能没有更多信息了")
        #     return Exception("加载信息失败，可能没有更多信息了")
        #     if(fastbreak!=None and fastbreak==StartText):
        #         if(needSetStartText):
        #             StartText=None 
        #     break
    tempMsg=wx.GetAllMessage(
            savepic   = False,   # 保存图片
            savefile  = False,   # 保存文件
            savevoice = False,    # 保存语音转文字内容
            saveVideo=False,
            saveZF=True,
            AreaText=(StartText,breakText)
        ) 
    return tempMsg 
def InsertPayDetail(toinsertInfo):
    global cursorsql
    insert_single_sql = '''INSERT INTO PayDetail (WXName, WXID ,AmountByXHS , AmountByUser  ,  ComputeDetail, DoCounts, ComputeDetailXHS ,DOIDS , Message  ,Remarks  ,Type ,HandleDate,InsertDate)
     VALUES (?, ?,?,?,?, ?,?,?,?, ?,?,?,?)'''      
    cursorsql.executemany(insert_single_sql, toinsertInfo)
    conn.commit()
class TimeoutException(Exception):
    pass

def call_with_timeout(func, args=(), kwargs={}, timeout=5):
    result = None
    exception = None
    
    def target():
        nonlocal result, exception
        try:
            result = func(*args, **kwargs)
        except Exception as e:
            exception = e
    
    thread = threading.Thread(target=target)
    thread.start()
    thread.join(timeout)
    
    if thread.is_alive():
        raise TimeoutException("Function call timed out")
    if exception:
        raise exception
    return result             
def GetNodeTextInfo(parameter):
    select_sql = "SELECT id,note_id,nickname,title,desc,time,likecount,collectedcount,commentcount,sharecount,image,xsec_token,user_id FROM NodeTextInfo WHERE note_id = ?" 
    if(UseInEncry==True):
        select_sql = "SELECT id,note_id,nickname,title,desc,time,likecount,collectedcount,commentcount,sharecount,image,xsec_token,user_id FROM NodeTextInfoEncry WHERE note_id = ?" 
    cursorsql.execute(select_sql, (parameter,))
    # 获取所有查询结果
    dataNode = cursorsql.fetchall()     
    if(UseInEncry==False):
        dataNode=[ [xor_encrypt_decrypt(datain) if isinstance(datain,int)==False else datain for datain in data ] for data in dataNode]
    return dataNode
def GetNodeTextInfo(parameter):
    select_sql = "SELECT id,note_id,nickname,title,desc,time,likecount,collectedcount,commentcount,sharecount,image,xsec_token,user_id FROM NodeTextInfo WHERE note_id = ?" 
    if(UseInEncry==True):
        select_sql = "SELECT id,note_id,nickname,title,desc,time,likecount,collectedcount,commentcount,sharecount,image,xsec_token,user_id FROM NodeTextInfoEncry WHERE note_id = ?" 
        parameter=xor_encrypt_decrypt(parameter)
    cursorsql.execute(select_sql, (parameter,))
    # 获取所有查询结果
    dataNode = cursorsql.fetchall()     
    if(UseInEncry==True):
        dataNode=[ [xor_encrypt_decrypt(datain) if isinstance(datain,int)==False else datain for datain in data ] for data in dataNode]
    return dataNode
def GetNodeHandleInfo():
    select_sql = "SELECT ID,noteID,handleUserID,handleUserName,handleUserImage,handleType,handleTime,mentionContent,status,addtime,fans FROM NodeHandleInfo " #where noteID='6857aa580000000012033ce3' and handleType in ('收藏')
    if(UseInEncry==True):
        select_sql = "SELECT ID,noteID,handleUserID,handleUserName,handleUserImage,handleType,handleTime,mentionContent,status,addtime,fans FROM NodeHandleInfoEncry " 
    cursorsql.execute(select_sql)
    # 获取所有查询结果
    dataNode = cursorsql.fetchall()    
    if(UseInEncry==True):  
        dataNode=[ [xor_encrypt_decrypt(datain) if isinstance(datain,int)==False else datain for datain in data ] for data in dataNode]
    return dataNode
if __name__ == '__main__':
    try:  
        global cursorsql,sht,sht1,wb,sht3,DZDay,StartText,breakText,noteToCalDetail,noteToCal
        UseInEncry=False
        file_path='config\\config.csv'
        dataread=[]
        endtimes=[]
        if os.path.exists(file_path): 
            with open(file_path, mode='r', newline='', encoding='utf-8') as file:
                reader = csv.reader(file)
                dataread = list(reader) 
                timetohandle=dataread[5]
                noteToCalDetail=dataread[2]
                noteToCal=dataread[1]
                if(len(dataread)<5 or timetohandle[0]=="" or (timetohandle[0]!="" and datetime.date.today()< datetime.datetime.strptime(timetohandle[0], "%Y/%m/%d").date())):
                    endtimes.append(datetime.date.today()- datetime.timedelta(days=1)) 
                else:
                    endtimes=[datetime.datetime.strptime(datadate, "%Y/%m/%d").date() for datadate in timetohandle if datadate!=""]

        StartText=None#"昨天 21:15"# "2025年5月30日 3:14"#"0:25"#"2025年4月25日 5:48" 如果是None会根据配置自动找到开始统计的地方
        breakText=None#"0:12"#"星期二 17:00"#"昨天 9:10" #None#终止查询的时间节点6:44 如果是None会根据配置自动找到结束统计的地方
        DZDay=endtimes#点赞收藏的哪天
        priceZ=1
        priceC=0.5
        priceP=0.5    
        wx = WeChat()
       #wx.ChatWith(who="姜可艾群")
        conn = sqlite3.connect('config\\WorkData.db')
        cursorsql = conn.cursor()
        CreateTableWxInfo()
        CreateTableWxToXHS() 
#---------------------------------------------------------获取聊天窗口控件信息--------------------------------------------------------------------
        msgsZF=[]
        msgsZF.extend(LoadFromZFMulty())
        msgsC =[]
        msgsC.extend(loadMoreCleaver((DZDay[0],DZDay[-1])))#-----------------向上滚动到开始统计的日期--------------
        msgs=[]
        msgs.extend(msgsC)
        msgs.extend(msgsZF)
#---------------------------------------------------------整理控件的数据到要保存的列表--------------------------------------------------------------------
        
        infosToSave=[]#从聊天记录中整理出来要保存的数据
        starthandle=False
        # 输出消息内容
        for msg in msgs:
            if(msg.content==StartText or len(msgsC)==0 or StartText==""):
                starthandle=True
            if(starthandle==False):
                continue
            if msg.type == 'sys':
                print(f'【系统消息】{msg.content}')
            elif msg.type == 'friend' or msg.type == 'text':
                if(breakText!=None and msg.content==breakText):break
                handletime=""
                if hasattr(msg, "time"):
                    handletime=msg.time
                sender = msg.sender # 这里可以将msg.sender改为msg.sender_remark，获取备注名
#---------------------------------------这些数据类型不处理，直接跳过----------------------------------------------------
                if(msg.myType=="[图片]" or msg.myType=="[文件]" or msg.myType=="[语音]" or msg.myType=="[已收款]"
                   or msg.sender=="Bb" or msg.sender=="馨"   ):
                    continue
                if(msg.myType=="[聊天记录]"): 
#----------------------------------------对转发的聊天记录进行处理--------------------------------------------------------
                    infossaveeone={}
                    infossaveeone["wxID"]=msg.sender
                    infossaveeone["zwxID"]=""#转发过来聊天记录里面的微信ID
                    infossaveeone["xhsID"]=""
                    infossaveeone["ZhengMing"]=""
                    infossaveeone["IsZ"]=0
                    infossaveeone["IsC"]=0
                    infossaveeone["IsP"]=0
                    infossaveeone["IsConfirm"]=0
                    infossaveeone["IsPay"]=0
                    infossaveeone["myType"]="[聊天记录]" 
                    infossaveeone["handletime"]=handletime
                    msgstring=""
                    for tempmsg in msg.content: 
                        if("msg" in tempmsg):
                            msgstring=msgstring+" "+tempmsg["content"]
                        else:
                            msgstring=msg.content
                            break
                    infossaveeone["content"]=msgstring

                    infosToSave.append(infossaveeone)
                else: 
#----------------------------------------对聊天记录进行处理--------------------------------------------------------
                    infossaveeone={}
                    find=False
                    xhsID=msg.content.replace("@姜可艾 没有结算完","").replace("赞","").replace("藏","").strip()
                    contentAfterHandle= msg.content.split("\n引用  的消息 :")[0]#content会包含引用的信息，引用信息有赞藏导致出错
#----------------------------------------为了把视频与文本记录对应，先查是否已经有存在的记录---------------------------
                    if(msg.myType=="[视频]"):
                        for infosave in reversed(infosToSave): 
                            if(infosave["wxID"]==sender and infosave["ZhengMing"]==""):
                                infossaveeone=infosave
                                find=True                        
                                break
                    elif(contentAfterHandle.find("@姜可艾 没有结算完")>-1):
                        for infosave in infosToSave:
                            if((infosave["wxID"]==sender and infosave["ZhengMing"]!="" and infosave["xhsID"]=="")):
                                    infossaveeone=infosave
                                    find=True
                                    break
                    else:
                        continue
#----------------------------------------没有的话就生成一条--------------------------------------------------------
                    if(find==False):
                        infossaveeone["wxID"]=msg.sender
                        infossaveeone["zwxID"]=""#转发过来聊天记录里面的微信ID
                        infossaveeone["xhsID"]=""
                        infossaveeone["ZhengMing"]=""#就是视频的地址，如果不保存视频，就是“[视频]”这几个字
                        infossaveeone["IsZ"]=0
                        infossaveeone["IsC"]=0
                        infossaveeone["IsP"]=0
                        infossaveeone["IsConfirm"]=0
                        infossaveeone["IsPay"]=0
                        infossaveeone["content"]=""
                        infossaveeone["myType"]="[聊天]"
                        infossaveeone["handletime"]=handletime
#----------------------------------------根据聊天记录类型对搜到的或者生成的数据进行赋值--------------------------------
                    if(msg.myType=="[视频]"): 
                        infossaveeone["ZhengMing"]=msg.content
                    else:  
                        if(contentAfterHandle.find("@姜可艾 没有结算完")>-1):
                            cs=1 #乘数
                            if(contentAfterHandle.find("2组赞藏")>-1 or contentAfterHandle.find("两组赞藏")>-1 or contentAfterHandle.find("两组")>-1 or contentAfterHandle.find("2组")>-1):
                                cs=2
                            if(contentAfterHandle.find("关注")>-1):
                                cs=1
                            if(contentAfterHandle.find("赞")>-1):
                                infossaveeone["IsZ"]=1*cs
                            if(contentAfterHandle.find("藏")>-1):
                                infossaveeone["IsC"]=1*cs
                            if(contentAfterHandle.find("评")>-1):
                                infossaveeone["IsP"]=1*cs
                            if(max([infossaveeone["IsZ"],infossaveeone["IsC"],infossaveeone["IsP"]])>=1):
                                infossaveeone["xhsID"]=xhsID
                            infossaveeone["content"]=msg.content
                    if(find==False):
                        infosToSave.append(infossaveeone)
                    #print(f'{sender.rjust(20)}：{msg.content}')

            elif msg.type == 'self':
                print(f'{msg.sender.ljust(20)}：{msg.content}')
            
            elif msg.type == 'time':
                if(breakText!=None and msg.content==breakText):break
                pass#print(f'\n【时间消息】{msg.time}')

            elif msg.type == 'recall':
                pass
                #print(f'【撤回消息】{msg.content}')
#---------------------------------------------------------找PayCOde和MarkID-----------------------------------------------------------------------------
        select_sql = "SELECT ID,wxName,MarkID,PayCode,OriginName,AddTime,Status,NeedPay FROM MarkWX" 
        cursorsql.execute(select_sql)
        # 获取所有查询结果
        dataNode2 = cursorsql.fetchall()   
        #先找数据库中originName完全一样的
        for info in infosToSave:
            info ["MarkID"]=0   
            info ["PayCode"]=0   
            for dn2 in dataNode2:
                    if(dn2[4] == info["wxID"] ):
                        info ["MarkID"]=dn2[2]
                        info ["PayCode"]=dn2[2] 
                        break 
        #先找微信名与数据库完全一样的，因为有的微信名包含在别人微信名里  
        for info in infosToSave:
            if(info ["MarkID"]!=0):
                continue
            for dn2 in dataNode2:
                    if(dn2[1] == info["wxID"] ):
                        print(f"{info['wxID']},{dn2[2]}的originName不一致")
                        info ["MarkID"]=dn2[2]
                        info ["PayCode"]=dn2[2] 
                        break      
        #再找互相包含的    
        for info in infosToSave:
            if(info ["MarkID"]!=0):
                continue
            for dn2 in dataNode2:
                    if((dn2[1] in info["wxID"] or info["wxID"] in dn2[1])):
                        print(f"{info['wxID']},{dn2[2]}的originName和wxName都不一致")
                        info ["MarkID"]=dn2[2]
                        info ["PayCode"]=dn2[2] 
                        break  
        #还是没有找到则插入新微信名到MarkWX表
        MarkID=dataNode2[-1][0]  #最后一个微信号的ID
        insertedMarkID=[]
        for info in infosToSave:#bug 有多个要插入的时候，最后一个ID为0了，要修复
            if(info["MarkID"]!=0 or info["wxID"] in insertedMarkID):
                continue
            MarkID=InsertMarkID(info["wxID"],dataNode2,MarkID)
            if MarkID !=None:
                info["MarkID"]=MarkID 
                info ["PayCode"]=MarkID 
                insertedMarkID.append(info["wxID"])
#---------------------------获取真实点赞的数据---------------------------
        
        dataNodeDZ1 = GetNodeHandleInfo() # 获取所有查询结果
        dataNodeDZ1Failed=[]#status为0的，但是已经点上的数据，为0是因为不符合要求，如只要10个，那第11个就是0，不接受了
        ZListConfirm=[]#小红书里面实际点赞成功的
        CListConfirm=[]#小红书里面实际收藏成功的
        OtherListConfirm=[]#小红书里面实际评论成功的

        #----------------------------------------将获取到的实际点赞情况根据config再次变更，只是变更读取的内存而不改数据库中的status-------------------------------------------------
        #由于可能有的要二次支付，如msc改为支付2元一个赞，所以需要第二次修改config，单独找出msc做的，再支付一遍，而calltest插入数据库中，可能还包含了其他的篇的status也为1
        for  DZ1 in dataNodeDZ1: 
            newstatus=DZ1[8]
            if(DZ1[8]==1 ):
                if(DZ1[1] in noteToCal):
                    # if(DZ1[10]==0): 
                    #     print(f'！！！！{wxname} 做{xhsid[5]}的账号 {xhsid[3]} 没有关注')
                    for i,ele in enumerate(noteToCal):
                        if DZ1[1]==ele: 
                            try:
                                if  DZ1[5]=="赞":
                                    if noteToCalDetail[i][0]!="1":
                                        newstatus=0 
                                elif DZ1[5]=="收藏":
                                    if noteToCalDetail[i][1]!="1":
                                        newstatus=0 
                                elif DZ1[5]=="评论":
                                    if noteToCalDetail[i][2]!="1":
                                        newstatus=0 
                            except Exception as ex:
                                print(ex)
                                traceback.print_exc()                        
                        if(newstatus==0):    
                            print(f'******{DZ1[2]} 对篇{DZ1[1]} 操作 {DZ1[5]} 不和要求')
                            break
                else:
                    continue 

            time1=datetime.datetime.strptime(DZ1[6], "%Y-%m-%d %H:%M:%S") 
            if(time1.date() in DZDay):
                if(newstatus==0):
                    dataNodeDZ1Failed.append(DZ1)
                    continue
                if(DZ1[5]=="赞"): 
                    ZListConfirm.append(DZ1)
                elif(DZ1[5]=="收藏"):
                    CListConfirm.append(DZ1)
                else:
                    OtherListConfirm.append(DZ1)

#----------------------------------------将转发的信息加入到@姜可艾2组的信息中，用来检查是否点成功了-------------------------------------------------
        touseZF=[]#转发的记录文字
        touseMiss=[]#转发的文字只有两组，2组，没有小红书名字，小红书名字在转发里
        for infosToSave1 in infosToSave:
            infosToSave1["contentAll"]=infosToSave1["content"]#转发记录+此信息内容
            if(infosToSave1["myType"]== "[聊天记录]"): 
                findob=None
                for tm in touseMiss:
                    if(tm["wxID"]==infosToSave1["wxID"] and tm["xhsID"] in("2组","两组")):
                        tm["contentAll"]+=infosToSave1["content"]
                        findob=tm 
                        break
                if(findob==None):
                    touseZF.append(infosToSave1)
                else:
                    touseMiss.remove(findob)
            if(infosToSave1["myType"]== "[聊天]"): 
                findob=None
                for tm in touseZF:
                    if(tm["wxID"]==infosToSave1["wxID"] ):
                        infosToSave1["contentAll"]+=tm["content"]
                        findob=tm 
                        break
                if(findob==None):
                    touseMiss.append(infosToSave1)
                else:
                    touseZF.remove(findob)        
#---------------------------------------------------------组装各个要插入的数据表-----------------------------------------------------------------------------
        toInsertSqlliteWXXHS=[]
        toInsertSqllite=[]
        toInsertXML=[]
        NotReceiveZC=[]
        ReceiveZC={} 
        CountSummary={"z":0,"c":0,"p":0,"Nz":0,"Nc":0,"Np":0,}#微信上统计的赞藏评个数，和没收到的赞藏评个数 ;还有 微信号对应的小红书号，内容+微信号 对应的聊天记录
        sortedInfosToSave = sorted(infosToSave, key=lambda p: p ["MarkID"])#先按MarkID排序，方便后面插入数据
      
        for infosToSave1 in sortedInfosToSave:
            CountSummary["z"]+=infosToSave1["IsZ"]
            CountSummary["c"]+=infosToSave1["IsC"]
            CountSummary["p"]+=infosToSave1["IsP"]  

            # if(infosToSave1["IsZ"]==0 and infosToSave1["IsC"]==0 and infosToSave1["IsP"]==0):
            #     continue


            payAmount=infosToSave1["IsZ"]*priceZ+infosToSave1["IsC"]*priceC+infosToSave1["IsP"]*priceP
            payAmountJS=payAmount#计算减去没在小红书查到的，有的人发 2组赞藏，不写小红书名字，查不到 #会写入excel表格中的“按用户发的然后从小红书查找计算”
            wid=   infosToSave1["contentAll"].replace("@姜可艾 没有结算完","").replace("赞",",").replace("藏",",").replace("\u2005",",").replace("。",",").replace("（）",",")\
                .replace("评",",").lower().replace("两组",",").replace("两组赞藏",",").replace("2组",",").replace("2组赞藏",",").replace("，",",").replace("\n",",").replace(" ",",")\
                .replace('[聊天记录]',",").replace("、","").replace("已自查",",").replace("）",",").replace("（",",").split("引用,,的消息")[0]#.replace(".",",")

            widl=[]#微信里面用户发的自己的小红书号
            for i in  wid.split(","):
                ddtt=i
                if(ddtt!=""):
                    ddtt=remove_chars_around_colon(i)
                if(ddtt!=""):
                    widl.append(ddtt)
            
            if (infosToSave1["wxID"] in CountSummary):
                CountSummary[infosToSave1["wxID"]]+=(","+wid)
                CountSummary["内容"+infosToSave1["wxID"]]+=infosToSave1["contentAll"]+infosToSave1["handletime"]+"\n\n"
            else:
                CountSummary[infosToSave1["wxID"]]=wid
                CountSummary["列表"+infosToSave1["wxID"]]=[]
                CountSummary["内容"+infosToSave1["wxID"]]=infosToSave1["contentAll"]+infosToSave1["handletime"]
            CountSummary["列表"+infosToSave1["wxID"]].extend(widl)


            payLoad=f"{','.join(widl)}： {str(infosToSave1['IsZ'])}*{str(priceZ)}+{str(infosToSave1['IsC'])}*{str(priceC)}+{str(infosToSave1['IsP'])}*{str(priceP)}\n"
            findedxhs=[]
#-------------------------------------------------------------------从获取的小红书数据中确认微信发的有没有收到---------------------------------------------------
            if(infosToSave1["IsZ"]>0):
                dataNodeDZ1FailedXHSID=  GetXHSID(dataNodeDZ1Failed,"赞")
                for zdata in ZListConfirm:
                    ziD=zdata[3].replace(" ", "").replace("，","").replace(" ","").lower()#从小红书里面取出的小红书号
                    if(ziD in widl  or ziD.replace("小红薯","")   in widl or ziD.replace("用户","") in widl):#if(ziD in wid or wid in ziD): 
                        findedxhs.append(ziD) 
                        if(infosToSave1["wxID"] in ReceiveZC ): 
                            if(zdata not in ReceiveZC[infosToSave1["wxID"]] ): 
                                ReceiveZC[infosToSave1["wxID"]].append(zdata) 
                        else:
                            ReceiveZC[infosToSave1["wxID"]]= [zdata]
                        if(zdata[10]==0):
                            CountSummary["Nz"]+=1
                            payAmountJS-=1*priceZ 
                            payLoad+=f"\n——{ziD}:Z{str(priceZ)}  "
                            remark="未关注" 
                            NotReceiveZC.append((infosToSave1["MarkID"],infosToSave1["wxID"],ziD,"赞",remark,infosToSave1["content"]))
                for ddd in widl: #看看用户发的小红书号是不是不在里面，好知道用户发的没有点上
                    if(ddd not in findedxhs and "小红薯"+ddd not in findedxhs and "用户"+ddd not in findedxhs):
                        CountSummary["Nz"]+=1
                        payAmountJS-=1*priceZ
                        payLoad+=f"\n——{ddd}:Z{str(priceZ)}  "
                        remark="未收到"
                        if(ddd in dataNodeDZ1FailedXHSID):
                            remark="收到,但不符合要求"
                        NotReceiveZC.append((infosToSave1["MarkID"],infosToSave1["wxID"],ddd,"赞",remark,infosToSave1["content"]))
                # for zii in ZListConfirm:
                #     if(zii[10]==0):
                #         CountSummary["Nz"]+=1
                #         payAmountJS-=1*priceZ
                #         payLoad+=f"\n——{ddd}:Z{str(priceZ)}  "
                #         remark="未关注" 
                #         NotReceiveZC.append((infosToSave1["MarkID"],infosToSave1["wxID"],ddd,"赞",remark,infosToSave1["content"]))
            findedxhs.clear()
            if(infosToSave1["IsC"]>0): 
                dataNodeDZ1FailedXHSID=  GetXHSID(dataNodeDZ1Failed,"收藏")
                for zdata in CListConfirm:
                    ziD=zdata[3].replace(" ", "").lower()
                    if(ziD in widl or ziD.replace("小红薯","")   in widl or ziD.replace("用户","") in widl):#if(ziD in wid or wid in ziD): 
                        findedxhs.append(ziD) 
                        if(infosToSave1["wxID"] in ReceiveZC ): 
                            if(zdata not in ReceiveZC[infosToSave1["wxID"]]): 
                                ReceiveZC[infosToSave1["wxID"]].append(zdata) 
                        else:
                            ReceiveZC[infosToSave1["wxID"]]= [zdata] 
                    if(zdata[10]==0):
                        CountSummary["Nc"]+=1
                        payAmountJS-=1*priceC
                        payLoad+=f"\n——{ziD}:Z{str(priceC)}  "
                        remark="未关注" 
                        NotReceiveZC.append((infosToSave1["MarkID"],infosToSave1["wxID"],ziD,"藏",remark,infosToSave1["content"]))
                for ddd in widl: 
                    if(ddd not in findedxhs and "小红薯"+ddd not in findedxhs and "用户"+ddd not in findedxhs):
                        CountSummary["Nc"]+=1
                        payAmountJS-=1*priceC
                        payLoad+=f"\n——{ddd}:C{str(priceC)}  "
                        remark="未收到"
                        if(ddd in dataNodeDZ1FailedXHSID):
                            remark="收到,但不符合要求"
                        NotReceiveZC.append((infosToSave1["MarkID"],infosToSave1["wxID"],ddd,"藏",remark,infosToSave1["content"]))
                # for zii in CListConfirm:
                #     if(zii[10]==0):
                #         CountSummary["Nc"]+=1
                #         payAmountJS-=1*priceC
                #         payLoad+=f"\n——{ddd}:Z{str(priceC)}  "
                #         remark="未关注" 
                #         NotReceiveZC.append((infosToSave1["MarkID"],infosToSave1["wxID"],ddd,"藏",remark,infosToSave1["content"]))
            findedxhs.clear()
            if(infosToSave1["IsP"]>0):
                dataNodeDZ1FailedXHSID=  GetXHSID(dataNodeDZ1Failed,"评论")
                for zdata in OtherListConfirm:
                    ziD=zdata[3].replace(" ", "").lower()
                    if(ziD in widl or ziD.replace("小红薯","") in widl or ziD.replace("用户","") in widl):#if(ziD in wid or wid in ziD): 
                        findedxhs.append(ziD) 
                        if(infosToSave1["wxID"] in ReceiveZC ): 
                            if(zdata not in ReceiveZC[infosToSave1["wxID"]]): 
                                ReceiveZC[infosToSave1["wxID"]].append(zdata) 
                        else:
                            ReceiveZC[infosToSave1["wxID"]]=[zdata] 
                    if(zdata[10]==0):
                        CountSummary["Np"]+=1
                        payAmountJS-=1*priceP
                        payLoad+=f"\n——{ziD}:Z{str(priceP)}  "
                        remark="未关注" 
                        NotReceiveZC.append((infosToSave1["MarkID"],infosToSave1["wxID"],ziD,"评",remark,infosToSave1["content"]))                            
                for ddd in widl: 
                    if(ddd not in findedxhs and "小红薯"+ddd not in findedxhs and "用户"+ddd not in findedxhs):
                        CountSummary["Np"]+=1
                        payAmountJS-=1*priceP
                        payLoad+=f"\n——{ddd}:P{str(priceP)}"
                        remark="未收到"
                        if(ddd in dataNodeDZ1FailedXHSID):
                            remark="收到,但不符合要求"
                        NotReceiveZC.append((infosToSave1["MarkID"],infosToSave1["wxID"],ddd,"评",remark,infosToSave1["content"])) 
                # for zii in OtherListConfirm:
                #     if(zii[10]==0):
                #         CountSummary["Np"]+=1
                #         payAmountJS-=1*priceP
                #         payLoad+=f"\n——{ddd}:Z{str(priceP)}  "
                #         remark="未关注" 
                #         NotReceiveZC.append((infosToSave1["MarkID"],infosToSave1["wxID"],ddd,"评",remark,infosToSave1["content"]))
#--------------------------------------------------------------------数据库列表的组装-------------------------------------------------------------                
            toInsertSqllite.append((infosToSave1["wxID"],infosToSave1["xhsID"],infosToSave1["IsZ"],infosToSave1["IsC"],infosToSave1["IsP"],infosToSave1["ZhengMing"],
                                    infosToSave1["IsConfirm"],infosToSave1["IsPay"],datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S"),infosToSave1["content"],infosToSave1["contentAll"]))
            if(infosToSave1["wxID"]!="姜可艾 没有结算完" and infosToSave1["wxID"]!="姜可艾" and infosToSave1["xhsID"]!=""):
                toInsertSqlliteWXXHS.append((infosToSave1["wxID"],infosToSave1["xhsID"],datetime.datetime.now().strftime("%Y_%m_%d_%H_%M_%S"),infosToSave1["MarkID"],infosToSave1["PayCode"])) 
#--------------------------------------------------------------------excel数据列表的组装----------------------------------------------------------------            
            find=False 
            maxC=max(infosToSave1['IsZ'],infosToSave1['IsC'],infosToSave1['IsP'])#操作了几个账号
            if(maxC>0):
                for info in toInsertXML:
                    if(infosToSave1["wxID"]==info[0]):
                        info[2]+=payAmount  #按用户发的总金额
                        info[4]=info[4]+("\n\n"+payLoad)  
                        info[5]+= maxC 
                        info[8]+=payAmountJS #按用户发的然后去小红书确认后的金额
                        find=True
                        break
                if(find==False ): 
                    toInsertXML.append([infosToSave1["wxID"],infosToSave1["MarkID"],payAmount,infosToSave1["PayCode"],payLoad,maxC, "",0,payAmountJS,"",""])
#-------------------------------------------------------------------------------对要插入excel的数据列表进行计算----------------------------------------------------------
        toinsertPayDetail=[]
        for wxname in ReceiveZC.keys():
            for xhsid in ReceiveZC[wxname]:
                resultCF = [key for key, value in ReceiveZC.items() if key!=wxname and any(xhsid[3] in item for item in value)]#多个人发了同一个小红书号
                if(len(resultCF)>0):
                    print(f'！！！！{wxname} 与 {",".join(resultCF)} 重复了 {xhsid[3]}')
                if(xhsid[10]==0):
                    print(f'！！！！{wxname} 做{xhsid[5]}的账号 {xhsid[3]} 没有关注，麻烦关注一下，再做数据')
        
        
        for txml in toInsertXML: 
            if(txml[0] in ReceiveZC) : 
                zlist=list(filter(lambda x: "赞" in x, ReceiveZC[txml[0]]))
                clist=list(filter(lambda x: "收藏" in x, ReceiveZC[txml[0]]))
                plist=list(filter(lambda x: "评论" in x, ReceiveZC[txml[0]]))
                zs=',\n'.join(f"{i[1]}({i[3]})" for i in zlist)
                cs=',\n'.join(f"{i[1]}({i[3]})" for i in clist)
                txml[6] =f"ActualZ:{str(len(zlist))} C:{str(len(clist))} P:{str(len(plist))}\n 赞:\n{zs}\n\n 藏:\n{cs}\n\n 评:{','.join(i[1]+i[3] for i in plist)} "
                txml[7] =len(zlist)*priceZ+len(clist)*priceC+len(plist)*priceP #按小红书查到的金额
            txml[9]=CountSummary[txml[0]] +"\n\n\n拆分后：\n\n"+ "\n".join(CountSummary["列表"+txml[0]])#小红书的账号
            txml[10]=CountSummary["内容"+txml[0]] #用户发的信息
            payDetailRemark=""
            if(txml[3]==0):
                payDetailRemark="无支付码"
            toinsertPayDetail.append([txml[0],txml[1],txml[7],txml[8],txml[4],txml[5],txml[6],txml[9],txml[10],payDetailRemark,"支付情况",f'{min(DZDay).strftime("%d")}-{max(DZDay).strftime("%d")}', datetime.datetime.today()])
#--------------------------------------------------------------------------看每一篇做了多少数据-------------------------------        
        ReceiveZC2= [value for d in ReceiveZC for value in ReceiveZC[d]]
        ReceiveZC2grouped = {}
        for rzc in ReceiveZC2:
            if(rzc[1] not in ReceiveZC2grouped): 
                # 获取所有查询结果
                dataNodeDZ1 = GetNodeTextInfo(rzc[1])   
                ReceiveZC2grouped[rzc[1]]=[]
                ReceiveZC2grouped[rzc[1]].extend([dataNodeDZ1,0,0,0])
            
            if(rzc[5]=="赞"):
                ReceiveZC2grouped[rzc[1]][1]+=1
            elif(rzc[5]=="收藏"):
                ReceiveZC2grouped[rzc[1]][2]+=1
            elif(rzc[5]=="评论"):
                ReceiveZC2grouped[rzc[1]][3]+=1
        for rzc in ReceiveZC2grouped:
            msgad=f"{(ReceiveZC2grouped[rzc][0][0][3]).ljust(40)}({ReceiveZC2grouped[rzc][1]}赞{ReceiveZC2grouped[rzc][2]}藏{ReceiveZC2grouped[rzc][3]}评)"
            print(msgad) 
            NotReceiveZC.append(("","",f" ",f" ",f"" ,msgad))
#-------------------------------------------------------------------------------对没有收到的赞藏评进行统计----------------------------------------------------------        
        
        NotReceiveZC.append((f"",f"",f"",f""
                             ,f""
                             ,f"微信【{str(CountSummary['z'])}，{str(CountSummary['c'])}，{str(CountSummary['p'])}】自然流量【{str(len(ZListConfirm)-(CountSummary['z']-CountSummary['Nz']))}，{str(len(CListConfirm)-(CountSummary['c']-CountSummary['Nc']))}，{str(len(OtherListConfirm)-(CountSummary['p']-CountSummary['Np']))}】"
                             f"总【{str(len(ZListConfirm))}，{str(len(CListConfirm))}，{str(len(OtherListConfirm))}】"))
        
        for nrzc in NotReceiveZC:
            toinsertPayDetail.append([nrzc[1],nrzc[0],0,0,"",0,nrzc[3],nrzc[2],nrzc[5],nrzc[4],"无赞藏",f'{min(DZDay).strftime("%d")}-{max(DZDay).strftime("%d")}', datetime.datetime.today()])
#-------------------------------------------------------------------------------将excel数据插入数据库----------------------------------------------------------


#---------------------------------------------------------插入数据库和Excel-----------------------------------------------------------------------------
        InsertWXToXHScache(toInsertSqlliteWXXHS)
        InsertWXInfoTocache(toInsertSqllite) 
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False    # 关闭一些提示信息，可以加快运行速度。 默认为 True。
        app.screen_updating = False    # 更新显示工作表的内容。默认为 True。关闭它也可以提升运行速度。
        wb = xw.Book()# app.books.open()# 
        sht = wb.sheets[0] 
        sht1 =wb.sheets.add(name='无支付码')
        sht3 =wb.sheets.add(name='无赞藏')
        InsertXMLNotReceive(NotReceiveZC)
        InsertXML(toInsertXML)
        InsertPayDetail(toinsertPayDetail)
    except Exception as ex:
        print(ex)
        traceback.print_exc()
    finally:
        # 关闭游标
        cursorsql.close()
        # 关闭数据库连接
        conn.close()
         