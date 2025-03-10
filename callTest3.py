# 从转发窗口获取信息，一般处理主群没有收到的信息的情况
from  wxautoMy.wxauto import uiautomation as uia
import time
from wxautoMy.wxauto  import WeChat
import xlwings as xw
import sqlite3 
import datetime
 
def LoadMoreMessage( C_MsgList: uia.ListControl):
        """加载当前聊天页面更多聊天信息
        
        Returns:
            bool: 是否成功加载更多聊天信息
        """
        loadmore = C_MsgList.GetLastChildControl()
        loadmore_bottom = loadmore.BoundingRectangle.bottom
        bottom = C_MsgList.BoundingRectangle.bottom
        while True:
            if loadmore.BoundingRectangle.bottom < bottom or loadmore.Name == '':
                isload = True
                break
            else:
                C_MsgList.WheelDown(wheelTimes=5, waitTime=0.1)
                if loadmore.BoundingRectangle.bottom == loadmore_bottom:
                    isload = False
                    break
                else:
                    loadmore_bottom = loadmore.BoundingRectangle.bottom
        C_MsgList.WheelDown(wheelTimes=2, waitTime=0.1)
        return isload

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
    insert_single_sql = '''INSERT INTO WXToXHSInfo (wxID ,xhsID ,AddTime,MarkID)
     VALUES (?, ?,?,?)'''      
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
    sht.range('A1') .value=['微信用户名','备注号',"支付金额","支付码"] 
    i=2
    for info in infos: 
        sht.range(f'A{i}') .value=list(info)
        i+=1  
    
    wb.save(f'Result\\微信信息{datetime.datetime.now().strftime("%Y_%m_%d_%H_%M_%S")}.xls')
    wb.close()
    

priceZ=1
priceC=0.5
priceP=0.5    
wx = WeChat()
conn = sqlite3.connect('config\\WorkData.db')
global cursorsql,sht,wb
cursorsql = conn.cursor()
myconf=uia.WindowControl(ClassName='ChatRecordWnd', searchDepth=1)
ddsd2=myconf.GetChildren()
ds2=ddsd2[1].ListControl()
msgs = []
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
                    msgs.append({"type":"text","sender":textbox1.Name,"content":textbox3.Name,"time":textbox2.Name,"msg":MsgItem2.Name})
    canload= LoadMoreMessage(ds2)
    time.sleep(1)


    try:
        # 输出消息内容
        for msg in msgs: 
            sender = msg["sender"] # 这里可以将msg.sender改为msg.sender_remark，获取备注名
            content=msg["content"]
            myType=msg["type"]
            if(  sender=="Bb" or sender=="馨"):
                continue
            if(myType=="[聊天记录]"):
                pass
                infossaveeone={}
                infossaveeone["wxID"]=sender
                infossaveeone["zwxID"]=""#转发过来聊天记录里面的微信ID
                infossaveeone["xhsID"]=""
                infossaveeone["ZhengMing"]=""
                infossaveeone["IsZ"]=0
                infossaveeone["IsC"]=0
                infossaveeone["IsP"]=0
                infossaveeone["IsConfirm"]=0
                infossaveeone["IsPay"]=0
                msgstring=""
                for tempmsg in content:
                    msgstring=msgstring+"___"+tempmsg["msg"]
                infossaveeone["content"]=msgstring

                infosToSave.append(infossaveeone)
            else: 
                infossaveeone={}
                find=False
                xhsID=content.replace("@姜可艾 没有结算完","_").replace("赞","_").replace("藏","_")
                if(myType=="[视频]"):
                    for infosave in reversed(infosToSave): 
                        if(infosave["wxID"]==sender and infosave["ZhengMing"]==""):
                            infossaveeone=infosave
                            find=True                        
                            break
                elif(content.find("@姜可艾 没有结算完")>-1):
                    for infosave in infosToSave:
                        if((infosave["wxID"]==sender and infosave["ZhengMing"]!="" and infosave["xhsID"]=="")):
                                infossaveeone=infosave
                                find=True
                                break
                else:
                    continue
                if(find==False):
                    infossaveeone["wxID"]=sender
                    infossaveeone["zwxID"]=""#转发过来聊天记录里面的微信ID
                    infossaveeone["xhsID"]=""
                    infossaveeone["ZhengMing"]=""
                    infossaveeone["IsZ"]=0
                    infossaveeone["IsC"]=0
                    infossaveeone["IsP"]=0
                    infossaveeone["IsConfirm"]=0
                    infossaveeone["IsPay"]=0
                    infossaveeone["content"]=""
                if(myType=="[视频]"): 
                    infossaveeone["ZhengMing"]=content
                else:  
                    if(content.find("@姜可艾 没有结算完")>-1):
                        cs=1
                        if(content.find("2组赞藏")>-1 or content.find("两组赞藏")>-1):
                            cs=2
                        if(content.find("关注")>-1):
                            cs=0
                        if(content.find("赞")>-1):
                            infossaveeone["IsZ"]=1*cs
                        if(content.find("藏")>-1):
                            infossaveeone["IsC"]=1*cs
                        if(content.find("评")>-1):
                            infossaveeone["IsP"]=1*cs
                        infossaveeone["xhsID"]=   content.replace("@姜可艾 没有结算完","_").replace("赞","_").replace("藏","_") 
                        infossaveeone["content"]=content
                if(find==False):
                    infosToSave.append(infossaveeone)
                #print(f'{sender.rjust(20)}：{msg.content}')

 

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
            toInsertSqlliteWXXHS.append((infosToSave1["wxID"],infosToSave1["xhsID"],datetime.datetime.now().strftime("%Y_%m_%d_%H_%M_%S"),infosToSave1["MarkID"])) 
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
# import xlwings as xw

# import sqlite3 

# def InsertWXToXHScache(infos):
#     select_sql = "SELECT * FROM WXToXHSInfo"
#     # 执行查询语句
#     global cursorsql
#     cursorsql.execute(select_sql)
#     # 获取所有查询结果
#     dataNode = cursorsql.fetchall() 
#     toinsertInfo=[]
#     for info in infos:
#         find=False
#         for dn in dataNode:
#             if(str(dn[1]) in str(info[1]) or str(info[1]) in str(dn[1])):
#                 info.append(dn[0])
#                 find=True
#                 break
#         if(find==False):
#             info.append(0)
#     #toinsertInfo=[]            
#     # 定义插入单条数据的 SQL 语句datetime.datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
#     insert_single_sql = '''INSERT INTO MarkWX (MarkID ,wxName ,wxidInDB)
#      VALUES (?, ?,?)'''
#     # 插入单条数据
#     cursorsql.executemany(insert_single_sql, infos)

 
#     conn.commit()
# global cursorsql,sht,wb
# conn = sqlite3.connect('config\\WorkData.db')
# cursorsql = conn.cursor()
# app = xw.App(visible=False, add_book=False)
# app.display_alerts = False    # 关闭一些提示信息，可以加快运行速度。 默认为 True。
# app.screen_updating = True    # 更新显示工作表的内容。默认为 True。关闭它也可以提升运行速度。
# wb =app.books.open("20250217.xlsx")#  xw.Book()# 
# sht = wb.sheets[0] 
# range_values = sht.range('A1:B202').value
# for row in range_values:
#     print(row)
# InsertWXToXHScache(range_values)