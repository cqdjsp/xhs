# 主群处理
from wxautoMy.wxauto  import WeChat
from  wxautoMy.wxauto import uiautomation as uia
import xlwings as xw
import sqlite3 
import datetime
import os
import traceback
from PIL import Image
import time
import csv
class DHMsg:
    def __init__(self, type, myType,sender,content,time,fromw):
        self.type = type
        self.myType = myType
        self.sender=sender
        self.content=content
        self.time=time
        self.fromw=fromw 
def xor_encrypt_decrypt(text, key):
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
            insert_single_sql = '''INSERT INTO MarkWX (wxName ,MarkID ,PayCode,OriginName,AddTime)
            VALUES (?, ?,?,?,?)'''  
            cursorsql.execute(insert_single_sql, (name,MarkID+1,0,name,datetime.datetime.now().strftime("%Y_%m_%d_%H_%M_%S")))
            # 提交事务，将更改保存到数据库
            conn.commit() 
            return  cursorsql.lastrowid
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
    select_sql = "SELECT * FROM WXToXHSInfo"
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
    sht3.range('A1') .value=['备注号' ,'微信用户名','小红书名',"内容","缺失","备注"] 
    i=2
    sht3.range('D:D').column_width = 50
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
    header=['微信用户名','备注号' ,"按用户发的计算","支付码图","按小红书查到的计算","金额计算过程","操作账号个数","实际操作数","按用户发的然后从小红书查找计算","小红书账号","信息内容","按小红书查到的计算==按用户发的然后从小红书查找计算"] 
    sht.range('A1') .value=header
    sht1.range('A1') .value=header#没有支付码的人
    sht1.range('F:F').column_width = 40
    sht1.range('H:H').column_width = 40
    sht1.range('K:K').column_width = 40
    sht.range('F:F').column_width = 40
    sht.range('H:H').column_width = 40
    sht.range('K:K').column_width = 40
    sht.range('J:J').column_width = 15
    i=2
    sht1i=2
    nopaycode=""
    paycodeMarkID=""
    for info in infos: 
        path=f'config\\zfcode\\{int(info[3])}.jpg'
        insertList=list((info[0],info[1],info[2],info[3],info[7],info[4],info[5],info[6],info[8],info[9],info[10]))
        
        if(int(info[3])==0):
            nopaycode+=f"@{info[0]}"
            paycodeMarkID+=f",{info[1]}"
            sht1.range(f'E{sht1i}').api.WrapText = True
            sht1.range(f'A{sht1i}') .value=insertList
            sht1.range(f'L{sht1i}').formula = f'=E{sht1i}=I{sht1i}'
            sht1i+=1

        else: 
            if(os.path.exists(path)):
                sht.range(f'L{i}').formula = f'=E{i}=I{i}'
                sht.range(f'E{i}').api.WrapText = True
                sht.range(f'A{i}') .value=insertList
                filePath = os.path.join(os.getcwd(),path )
                add_center(sht, 'D'+str(i), filePath, width=350, height=350)
                i+=1  
            else:
                print(f"{info[0]},MarkID{info[1]}在数据库有paycode但是没有文件")
                sht1.range(f'E{sht1i}').api.WrapText = True
                sht1.range(f'A{sht1i}') .value=insertList
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
                                    if(mg.content==text_control.Name and mg.sender==textbox1.Name):
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
                                if(mg.content==textbox3.Name and mg.sender==textbox1.Name):
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
def check_window_exists(class_name,findindex=1):
    try:
        # 尝试查找具有指定 ClassName 的 WindowControl
        window = uia.WindowControl(ClassName=class_name, searchDepth=1,foundIndex=findindex)
        # 如果找到了控件，确保它是有效的（没有被销毁）
        if window.Exists():
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
def loadMoreCleaver(AreaText):
    startText=AreaText[0]
    tempMsg=[]
    while(True) :
        tempMsg=wx.GetAllMessage(
                savepic   = False,   # 保存图片
                savefile  = False,   # 保存文件
                savevoice = False,    # 保存语音转文字内容
                saveVideo=False,
                saveZF=False,
                AreaText=AreaText
            ) 
        if (startText!="" and  len([msg for msg in tempMsg if startText in msg.content])>0):
            break
        elif(wx.LoadMoreMessage()):
            print("向上滚动")   
        else:
            print("加载信息结束") 
            break
    tempMsg=wx.GetAllMessage(
            savepic   = False,   # 保存图片
            savefile  = False,   # 保存文件
            savevoice = False,    # 保存语音转文字内容
            saveVideo=False,
            saveZF=True,
            AreaText=AreaText
        ) 
    return tempMsg 
if __name__ == '__main__':
    try:  
        global cursorsql,sht,sht1,wb,sht3,DZDay
        file_path='config\\config.csv'
        dataread=[]
        if os.path.exists(file_path): 
            with open(file_path, mode='r', newline='', encoding='utf-8') as file:
                reader = csv.reader(file)
                dataread = list(reader)
                cookie = dataread[0][0]
                noteToCal=dataread[1]
                endtimes=[datetime.datetime.strptime(datadate, "%Y/%m/%d").date() for datadate in dataread[2] if datadate!=""]
                catchlike= int(dataread[3][0])
                catchMention=int (dataread[3][1])  
        IsZF=True#是否是从转发的窗口获取数据
        StartText="昨天 1:29"#"昨天 8:15"#"0:25"#"2025年4月25日 5:48"
        breakText="8:30"#"星期二 17:00"#"昨天 9:10" #None#终止查询的时间节点6:44
        DZDay=endtimes#点赞收藏的哪天
        priceZ=1
        priceC=0.5
        priceP=0.5    
        wx = WeChat()
        wx.ChatWith(who="姜可艾群")
        conn = sqlite3.connect('config\\WorkData.db')
        cursorsql = conn.cursor()
        CreateTableWxInfo()
        CreateTableWxToXHS() 
#---------------------------------------------------------获取聊天窗口控件信息--------------------------------------------------------------------
        msgsZF=[]
        msgsZF.extend(LoadFromZFMulty())
        msgsC =[]
        msgsC.extend(loadMoreCleaver((StartText,breakText)))
        msgs=[]
        msgs.extend(msgsC)
        msgs.extend(msgsZF)
#---------------------------------------------------------整理控件的数据到要保存的列表--------------------------------------------------------------------
        infosToSave=[]
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
                            cs=1
                            if(contentAfterHandle.find("2组赞藏")>-1 or contentAfterHandle.find("两组赞藏")>-1 or contentAfterHandle.find("两组")>-1 or contentAfterHandle.find("2组")>-1):
                                cs=2
                            if(contentAfterHandle.find("关注")>-1):
                                cs=0
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
        select_sql = "SELECT * FROM MarkWX" 
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
        for info in infosToSave:
            if(info["MarkID"]!=0):
                continue
            MarkID=InsertMarkID(info["wxID"],dataNode2,MarkID)
            if MarkID !=None:
                info["MarkID"]=MarkID 
#---------------------------获取真实点赞的数据---------------------------
        select_sql = "SELECT * FROM NodeHandleInfo" 
        cursorsql.execute(select_sql)
        # 获取所有查询结果
        dataNodeDZ1 = cursorsql.fetchall() 
        dataNodeDZ1Failed=[]#status为0的，但是已经点上的数据，为0是因为不符合要求，如只要10个，那第11个就是0，不接受了
        ZListConfirm=[]
        CListConfirm=[]
        OtherListConfirm=[]
        for  DZ1 in dataNodeDZ1:
            time1=datetime.datetime.strptime(DZ1[6], "%Y-%m-%d %H:%M:%S") 
            if(time1.date() in DZDay):
                if(DZ1[8]==0):
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
        for infosToSave1 in infosToSave:
            CountSummary["z"]+=infosToSave1["IsZ"]
            CountSummary["c"]+=infosToSave1["IsC"]
            CountSummary["p"]+=infosToSave1["IsP"]  
            payAmount=infosToSave1["IsZ"]*priceZ+infosToSave1["IsC"]*priceC+infosToSave1["IsP"]*priceP
            payAmountJS=payAmount#计算减去没在小红书查到的，有的人发 2组赞藏，不写小红书名字，查不到
            wid=   infosToSave1["contentAll"].replace("@姜可艾 没有结算完","").replace("赞",",").replace("藏",",").replace("\u2005",",").replace("。",",").replace("（）",",")\
                .replace("评",",").lower().replace("两组",",").replace("两组赞藏",",").replace("2组",",").replace("2组赞藏",",").replace("，",",").replace("\n",",").replace(" ",",")\
                .replace('[聊天记录]',",").replace("已自查",",").replace("）",",").replace("（",",").split("引用,,的消息")[0]#.replace(".",",")

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
                for ddd in widl: 
                    if(ddd not in findedxhs and "小红薯"+ddd not in findedxhs and "用户"+ddd not in findedxhs):
                        CountSummary["Nz"]+=1
                        payAmountJS-=1*priceZ
                        payLoad+=f"\n——{ddd}:Z{str(priceZ)}  "
                        remark=""
                        if(ddd in dataNodeDZ1FailedXHSID):
                            remark="收到,但不符合要求"
                        NotReceiveZC.append((infosToSave1["MarkID"],infosToSave1["wxID"],ddd,infosToSave1["content"],"赞",remark))
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
                for ddd in widl: 
                    if(ddd not in findedxhs and "小红薯"+ddd not in findedxhs and "用户"+ddd not in findedxhs):
                        CountSummary["Nc"]+=1
                        payAmountJS-=1*priceC
                        payLoad+=f"\n——{ddd}:C{str(priceC)}  "
                        remark=""
                        if(ddd in dataNodeDZ1FailedXHSID):
                            remark="收到,但不符合要求"
                        NotReceiveZC.append((infosToSave1["MarkID"],infosToSave1["wxID"],ddd,infosToSave1["content"],"藏",remark))
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
                for ddd in widl: 
                    if(ddd not in findedxhs and "小红薯"+ddd not in findedxhs and "用户"+ddd not in findedxhs):
                        CountSummary["Np"]+=1
                        payAmountJS-=1*priceP
                        payLoad+=f"\n——{ddd}:P{str(priceP)}"
                        remark=""
                        if(ddd in dataNodeDZ1FailedXHSID):
                            remark="收到,但不符合要求"
                        NotReceiveZC.append((infosToSave1["MarkID"],infosToSave1["wxID"],ddd,infosToSave1["content"],"评",remark)) 
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
                        info[2]+=payAmount
                        info[4]=info[4]+("\n\n"+payLoad)
                        info[5]+= maxC 
                        info[8]+=payAmountJS
                        find=True
                        break
                if(find==False ): 
                    toInsertXML.append([infosToSave1["wxID"],infosToSave1["MarkID"],payAmount,infosToSave1["PayCode"],payLoad,maxC, "",0,payAmountJS,"",""])
#-------------------------------------------------------------------------------对要插入excel的数据列表进行计算----------------------------------------------------------
        for txml in toInsertXML: 
            if(txml[0] in ReceiveZC) : 
                zlist=list(filter(lambda x: "赞" in x, ReceiveZC[txml[0]]))
                clist=list(filter(lambda x: "收藏" in x, ReceiveZC[txml[0]]))
                plist=list(filter(lambda x: "评论" in x, ReceiveZC[txml[0]]))
                zs=',\n'.join(f"{i[1]}({i[3]})" for i in zlist)
                cs=',\n'.join(f"{i[1]}({i[3]})" for i in clist)
                txml[6] =f"ActualZ:{str(len(zlist))} C:{str(len(clist))} P:{str(len(plist))}\n 赞:\n{zs}\n\n 藏:\n{cs}\n\n 评:{','.join(i[1]+i[3] for i in plist)} "
                txml[7] =len(zlist)*priceZ+len(clist)*priceC+len(plist)*priceP
            txml[9]=CountSummary[txml[0]] +"\n\n\n拆分后：\n\n"+ "\n".join(CountSummary["列表"+txml[0]])#小红书的账号
            txml[10]=CountSummary["内容"+txml[0]] #用户发的信息
        NotReceiveZC.append((f"微信赞{str(CountSummary['z'])}",f"微信藏{str(CountSummary['c'])}",f"微信评{str(CountSummary['p'])}",f"小红书赞{str(len(ZListConfirm))}藏{str(len(CListConfirm))}评{str(len(OtherListConfirm))}"
                             ,f"自然流量赞:{str(len(ZListConfirm)-(CountSummary['z']-CountSummary['Nz']))}藏:{str(len(CListConfirm)-(CountSummary['c']-CountSummary['Nc']))}评:{str(len(OtherListConfirm)-(CountSummary['p']-CountSummary['Np']))}"
                             ,""))
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
    except Exception as ex:
        print(ex)
        traceback.print_exc()
    finally:
        # 关闭游标
        cursorsql.close()
        # 关闭数据库连接
        conn.close()
         