# 主群处理
from wxautoMy.wxauto  import WeChat
import xlwings as xw
import sqlite3 
import datetime
import os
from PIL import Image
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
    insert_single_sql = '''INSERT INTO WXHandleInfo (wxID ,xhsID ,IsZ , IsC, IsP, ZhengMing , IsConfirm ,IsPay,addTime,msg)
     VALUES (?, ?,?,?, ?,?,?, ?,?, ?)'''
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
            toinsertInfo.append(info)

     
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
def InsertXML(infos):
    global sht,wb 
    wxdata=[{}]#[{"wx":"wx用户名"，"xhs":"xhs用户名","z":"1","C":"1","P":"0","sp":"视频证明地址"}]
    # 将a1,a2,a3输入第一列，b1,b2,b3输入第二列
    sht.range('A1') .value=['微信用户名','备注号' ,"支付金额","支付码图"] 
    i=2
    for info in infos: 
        sht.range(f'A{i}') .value=list((info[0],info[1],info[2]))
        path=f'config\\zfcode\\{int(info[3])}.jpg'
        if(os.path.exists(path)):
            filePath = os.path.join(os.getcwd(),path )
            add_center(sht, 'D'+str(i), filePath, width=350, height=350)
        i+=1  
    
    wb.save(f'Result\\结算{datetime.datetime.now().strftime("%Y_%m_%d_%H_%M_%S")}.xls')
    wb.close()
if __name__ == '__main__':
    try:  
        priceZ=1
        priceC=0.5
        priceP=0.5    
        wx = WeChat()
        conn = sqlite3.connect('config\\WorkData.db')
        global cursorsql,sht,wb
        cursorsql = conn.cursor()
        CreateTableWxInfo()
        CreateTableWxToXHS() 
        # 获取当前聊天窗口消息
        msgs = wx.GetAllMessage(
            savepic   = False,   # 保存图片
            savefile  = False,   # 保存文件
            savevoice = False,    # 保存语音转文字内容
            saveVideo=False,
            saveZF=True
        ) 

        infosToSave=[]
        # 输出消息内容
        for msg in msgs:
            if msg.type == 'sys':
                print(f'【系统消息】{msg.content}')
            elif msg.type == 'friend':
                sender = msg.sender # 这里可以将msg.sender改为msg.sender_remark，获取备注名
                if(msg.myType=="[图片]" or msg.myType=="[文件]" or msg.myType=="[语音]" or msg.myType=="[已收款]"
                   or msg.sender=="Bb" or msg.sender=="馨"):
                    continue
                if(msg.myType=="[聊天记录]"):
                    pass
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
                    msgstring=""
                    for tempmsg in msg.content:
                        msgstring=msgstring+"___"+tempmsg["msg"]
                    infossaveeone["content"]=msgstring

                    infosToSave.append(infossaveeone)
                else: 
                    infossaveeone={}
                    find=False
                    xhsID=msg.content.replace("@姜可艾 没有结算完","_").replace("赞","_").replace("藏","_")
                    if(msg.myType=="[视频]"):
                        for infosave in reversed(infosToSave): 
                            if(infosave["wxID"]==sender and infosave["ZhengMing"]==""):
                                infossaveeone=infosave
                                find=True                        
                                break
                    elif(msg.content.find("@姜可艾 没有结算完")>-1):
                        for infosave in infosToSave:
                            if((infosave["wxID"]==sender and infosave["ZhengMing"]!="" and infosave["xhsID"]=="")):
                                    infossaveeone=infosave
                                    find=True
                                    break
                    else:
                        continue
                    if(find==False):
                        infossaveeone["wxID"]=msg.sender
                        infossaveeone["zwxID"]=""#转发过来聊天记录里面的微信ID
                        infossaveeone["xhsID"]=""
                        infossaveeone["ZhengMing"]=""
                        infossaveeone["IsZ"]=0
                        infossaveeone["IsC"]=0
                        infossaveeone["IsP"]=0
                        infossaveeone["IsConfirm"]=0
                        infossaveeone["IsPay"]=0
                        infossaveeone["content"]=""
                    if(msg.myType=="[视频]"): 
                        infossaveeone["ZhengMing"]=msg.content
                    else:  
                        if(msg.content.find("@姜可艾 没有结算完")>-1):
                            cs=1
                            if(msg.content.find("2组赞藏")>-1 or msg.content.find("两组赞藏")>-1):
                                cs=2
                            if(msg.content.find("关注")>-1):
                                cs=0
                            if(msg.content.find("赞")>-1):
                                infossaveeone["IsZ"]=1*cs
                            if(msg.content.find("藏")>-1):
                                infossaveeone["IsC"]=1*cs
                            if(msg.content.find("评")>-1):
                                infossaveeone["IsP"]=1*cs
                            infossaveeone["xhsID"]=   msg.content.replace("@姜可艾 没有结算完","_").replace("赞","_").replace("藏","_") 
                            infossaveeone["content"]=msg.content
                    if(find==False):
                        infosToSave.append(infossaveeone)
                    #print(f'{sender.rjust(20)}：{msg.content}')

            elif msg.type == 'self':
                print(f'{msg.sender.ljust(20)}：{msg.content}')
            
            elif msg.type == 'time':
                pass#print(f'\n【时间消息】{msg.time}')

            elif msg.type == 'recall':
                print(f'【撤回消息】{msg.content}')

        select_sql = "SELECT * FROM MarkWX" 
        cursorsql.execute(select_sql)
        # 获取所有查询结果
        dataNode2 = cursorsql.fetchall()   
        #先找微信名与数据库完全一样的，因为有的微信名包含在别人微信名里  
        for info in infosToSave:
            info ["MarkID"]=0   
            info ["PayCode"]=0   
            for dn2 in dataNode2:
                    if(dn2[1] == info["wxID"] ):
                        info ["MarkID"]=dn2[2]
                        info ["PayCode"]=dn2[3] 
                        break          
        for info in infosToSave:
            if(info ["MarkID"]!=0):
                continue
            for dn2 in dataNode2:
                    if((dn2[1] in info["wxID"] or info["wxID"] in dn2[1])):
                        info ["MarkID"]=dn2[2]
                        info ["PayCode"]=dn2[3] 
                        break   

        toInsertSqlliteWXXHS=[]
        toInsertSqllite=[]
        toInsertXML=[]
        for infosToSave1 in infosToSave:
            toInsertSqllite.append((infosToSave1["wxID"],infosToSave1["xhsID"],infosToSave1["IsZ"],infosToSave1["IsC"],infosToSave1["IsP"],infosToSave1["ZhengMing"],
                                    infosToSave1["IsConfirm"],infosToSave1["IsPay"],datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S"),infosToSave1["content"]))
            if(infosToSave1["wxID"]!="姜可艾 没有结算完"):
                toInsertSqlliteWXXHS.append((infosToSave1["wxID"],infosToSave1["xhsID"],datetime.datetime.now().strftime("%Y_%m_%d_%H_%M_%S"),infosToSave1["MarkID"],infosToSave1["PayCode"])) 
            payAmount=infosToSave1["IsZ"]*priceZ+infosToSave1["IsC"]*priceC+infosToSave1["IsP"]*priceP
            
            find=False 
            for info in toInsertXML:
                if(infosToSave1["wxID"]==info[0]):
                    info[2]+=payAmount
                    find=True
                    break
            if(find==False):
                toInsertXML.append([infosToSave1["wxID"],infosToSave1["MarkID"],payAmount,infosToSave1["PayCode"]])
        InsertWXToXHScache(toInsertSqlliteWXXHS)
        InsertWXInfoTocache(toInsertSqllite) 
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False    # 关闭一些提示信息，可以加快运行速度。 默认为 True。
        app.screen_updating = True    # 更新显示工作表的内容。默认为 True。关闭它也可以提升运行速度。
        wb = xw.Book()# app.books.open()# 
        sht = wb.sheets[0] 
        InsertXML(toInsertXML)
    except Exception as ex:
        print(ex)
    finally:
        # 关闭游标
        cursorsql.close()
        # 关闭数据库连接
        conn.close()
         