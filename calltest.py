#获取小红书的赞藏评数据
import datetime 
from time import sleep
import csv
import random
import re
from playwright.sync_api import sync_playwright
import datetime
from xhs.core import DataFetchError, XhsClient
import os 
import sqlite3 
import traceback
 
def sign(uri, data=None, a1="", web_session=""):
    for _ in range(10):
        try:
            with sync_playwright() as playwright:
                stealth_js_path = "E:\\my\\job\\xhsTG\\public\\stealth.min.js"
                chromium = playwright.chromium
                browser_path = os.path.join(os.getenv('LOCALAPPDATA'), 'ms-playwright', 'chromium-1148', 'chrome-win', 'chrome.exe')
                # 如果一直失败可尝试设置成 False 让其打开浏览器，适当添加 sleep 可查看浏览器状态
                browser = chromium.launch(executable_path=browser_path,headless=True)

                browser_context = browser.new_context()
                browser_context.add_init_script(path=stealth_js_path)
                context_page = browser_context.new_page()
                context_page.goto("https://www.xiaohongshu.com")
                browser_context.add_cookies([
                    {'name': 'a1', 'value': a1, 'domain': ".xiaohongshu.com", 'path': "/"}]
                )
                context_page.reload()
                # 这个地方设置完浏览器 cookie 之后，如果这儿不 sleep 一下签名获取就失败了，如果经常失败请设置长一点试试
                sleep(1)
                encrypt_params = context_page.evaluate("([url, data]) => window._webmsxyw(url, data)", [uri, data])
                return {
                    "x-s": encrypt_params["X-s"],
                    "x-t": str(encrypt_params["X-t"])
                }
        except Exception as ex:
            print(ex)
            # 这儿有时会出现 window._webmsxyw is not a function 或未知跳转错误，因此加一个失败重试趴
            pass
    raise Exception("重试了这么多次还是无法签名成功，寄寄寄")
def remove_brackets_content(s):
    # 使用正则表达式匹配中括号及其中的内容
    result = re.sub(r'\[.*?\]', '', s)
    return result
#按顺序获取指定个数个获取点赞，收藏，评论
def GetInfoBySeq(catchlike,catchMention):
    global cursorsql,noteToCal,endtimes
    # 定义查询数据的 SQL 语句
    select_sql = "SELECT id,note_id,xsec_token,user_id,nickname,title,desc,time,likecount,collectedcount,commentcount,sharecount,image FROM NodeTextInfoEncry"
    if(InNormal):
        select_sql = "SELECT id,note_id,xsec_token,user_id,nickname,title,desc,time,likecount,collectedcount,commentcount,sharecount,image  FROM NodeTextInfo"
    # 执行查询语句
    cursorsql.execute(select_sql)
    # 获取所有查询结果
    dataNode = cursorsql.fetchall()
    if(InNormal==False):
        dataNode=[ [xor_encrypt_decrypt(datain) if isinstance(datain,int)==False else datain for datain in data ] for data in dataNode]
    data = [ ] 
    noteinfoTitle=""
    global handleType
    global xhs_client
    cursor=""
    minendtime=min(endtimes)
   
    while(catchlike>0): 
        note=  xhs_client.get_like_notifications(20,cursor)
 
        findendtime=False#找到停止的时间也不再找了。退出while
        incurrectDay=False
        for noteInfo in note['message_list']:
            incurrectDay=False
            for endtime in endtimes: 
                if datetime.datetime.fromtimestamp(noteInfo['time']).date()!=endtime:
                    if datetime.datetime.fromtimestamp(noteInfo['time']).date()<minendtime:
                        findendtime=True
                        break 
                    else:
                        continue  
                else:
                    incurrectDay=True
                    break
            
            if(findendtime):break
            if(incurrectDay==False):continue
            if( noteInfo['type'] in handleType):
                    noteID=noteInfo['item_info']['id'] if 'liked/item' ==noteInfo['type'] else noteInfo['item_info']["attach_item_info"]["id"] 
                    noteSec=noteInfo['item_info']['xsec_token'] if 'liked/item' ==noteInfo['type'] else noteInfo['item_info']["attach_item_info"]["xsec_token"] 
                    find=False
                    for notedatacache in dataNode:
                        if(notedatacache[1]==noteID):
                            noteinfoTitle=notedatacache[5]
                            find=True
                            break
                    if(find==False):
                        try:  
                            noteinfoSWeb=xhs_client.get_note_by_id(noteID,noteSec)
                        except Exception as ex:
                            print(ex)
                            traceback.print_exc()    
                            continue
                        InsertNoteInfoTocache(noteID,noteSec,noteinfoSWeb["user"]["user_id"],noteinfoSWeb["user"]["nickname"],noteinfoSWeb["title"],noteinfoSWeb["desc"],datetime.datetime.fromtimestamp(noteinfoSWeb['time']/1000).strftime("%Y-%m-%d %H:%M:%S"),
                                            (noteinfoSWeb["interact_info"]["liked_count"]),(noteinfoSWeb["interact_info"]["collected_count"]),(noteinfoSWeb["interact_info"]["comment_count"]),(noteinfoSWeb["interact_info"]["share_count"]),noteInfo['item_info']['image'])
                        noteinfoTitle=noteinfoSWeb["title"]
                        cursorsql.execute(select_sql)
                        # 获取所有查询结果
                        dataNode = cursorsql.fetchall()
                        if(InNormal==False):
                            dataNode=[ [xor_encrypt_decrypt(datain) if isinstance(datain,int)==False else datain for datain in data ] for data in dataNode]
                    data.append({"预览图":noteInfo['item_info']['image'],"篇":noteID,"篇title":noteinfoTitle,
                                "作者":noteInfo['item_info']['user_info']["userid"] if  'user_info' in noteInfo['item_info'] else "",
                                '操作人ID':noteInfo["user_info"]['userid'], '操作人昵称':noteInfo["user_info"]['nickname'],'操作人头像':noteInfo["user_info"]['image'],
                                '操作类型':handleType[noteInfo['type']],'操作时间':datetime.datetime.fromtimestamp(noteInfo['time']).strftime("%Y-%m-%d %H:%M:%S"),'价格':handleTypePrice[noteInfo['type']],
                                '评论内容': "",'操作ID': str(noteInfo['time']) +'_'+noteInfo['id'],'关注':noteInfo["user_info"]['indicator']=="你的粉丝" if 'indicator' in noteInfo["user_info"]  else False
                            }  )
        if(findendtime): break    
        catchlike-=20
        cursor=note['strCursor']  
        if(note["has_more"]==False):break  
        print(f"赞藏：catchlike{catchlike} cursor:{cursor}")
        sleep(random.uniform(0.1, 0.5))
    cursor=""
    while(catchMention>0):
        mentionNote=  xhs_client.get_mention_notifications(20,cursor)
        findendtime=False#找到停止的时间也不再找了。退出while
        for noteInfo in mentionNote['message_list'] :
            incurrectDay=False
            for endtime in endtimes: 
                if datetime.datetime.fromtimestamp(noteInfo['time']).date()!=endtime:
                    if datetime.datetime.fromtimestamp(noteInfo['time']).date()<minendtime:
                        findendtime=True
                        break 
                    else:
                        continue  
                else:
                    incurrectDay=True
                    break
            
            if(findendtime):break
            if(incurrectDay==False):continue
            if( noteInfo['type'] in handleType and 'target_comment' not in noteInfo['comment_info']  and remove_brackets_content(noteInfo['comment_info']["content"])!="" ):
                noteID=noteInfo['item_info']['id'] if 'id' in noteInfo['item_info'] else ""
                # if( noteID not in  noteToCal ):
                #     continue
            
                noteSec=noteInfo['item_info']['xsec_token'] if 'xsec_token' in noteInfo['item_info'] else ""
                find=False
                for notedatacache in dataNode:
                    if(notedatacache[1]==noteID):
                        noteinfoTitle=notedatacache[5]
                        find=True
                        break
                if(find==False): 
                    try:  
                        noteinfoSWeb=xhs_client.get_note_by_id(noteID,noteSec)
                    except Exception as ex:
                        print(ex)
                        traceback.print_exc()    
                        continue
                    InsertNoteInfoTocache(noteID,noteSec,noteinfoSWeb["user"]["user_id"],noteinfoSWeb["user"]["nickname"],noteinfoSWeb["title"],noteinfoSWeb["desc"],datetime.datetime.fromtimestamp(noteinfoSWeb['time']/1000).strftime("%Y-%m-%d %H:%M:%S"),
                                            int(noteinfoSWeb["interact_info"]["liked_count"]),int(noteinfoSWeb["interact_info"]["collected_count"]),int(noteinfoSWeb["interact_info"]["comment_count"]),int(noteinfoSWeb["interact_info"]["share_count"]),noteInfo['item_info']['image'])
                    noteinfoTitle=noteinfoSWeb["title"]
                    cursorsql.execute(select_sql)
                    # 获取所有查询结果
                    dataNode = cursorsql.fetchall()
                    if(InNormal==False):
                        dataNode=[ [xor_encrypt_decrypt(datain) if isinstance(datain,int)==False else datain for datain in data ] for data in dataNode]
                data.append({"预览图":noteInfo['item_info']['image'],"篇":noteID,"篇title":noteinfoTitle,
                            "作者":noteInfo['item_info']['user_info']["userid"] if  'user_info' in noteInfo['item_info'] else "",
                            '操作人ID':noteInfo["user_info"]['userid'], '操作人昵称':noteInfo["user_info"]['nickname'],'操作人头像':noteInfo["user_info"]['image'],
                            '操作类型':handleType[noteInfo['type']],'操作时间':datetime.datetime.fromtimestamp(noteInfo['time']).strftime("%Y-%m-%d %H:%M:%S"),'价格':handleTypePrice[noteInfo['type']],
                            '评论内容': noteInfo['comment_info']["content"],'操作ID': noteInfo['comment_info']["id"],'关注':noteInfo["user_info"]['indicator']=="你的粉丝" if 'indicator' in noteInfo["user_info"]  else False
                        }  )
        if(findendtime): break 
        catchMention-=20
        cursor=mentionNote['strCursor']  
        if(mentionNote["has_more"]==False):break  
        print(f"评论：mention{catchMention} cursor:{cursor}")
        sleep(random.uniform(0.2, 2.0))
   
    return data
def InsertNoteInfoTocacheEncry(note_id ,xsec_token ,user_id , nickname="", title="", desc="" , time="" ,likecount=0 ,collectedcount=0, commentcount=0 ,sharecount=0,image=""):
    global cursorsql
    # 定义插入单条数据的 SQL 语句
    insert_single_sql = '''INSERT INTO NodeTextInfoEncry (note_id ,xsec_token ,user_id , nickname, title, desc , time ,likecount ,collectedcount, commentcount ,sharecount,image)
     VALUES (?, ?,?,?, ?,?,?, ?,?,?, ?,?)'''
    # 插入单条数据
    cursorsql.execute(insert_single_sql, (xor_encrypt_decrypt(note_id) ,xor_encrypt_decrypt(xsec_token)  ,xor_encrypt_decrypt(user_id ) ,xor_encrypt_decrypt( nickname) ,xor_encrypt_decrypt( title),
                                        xor_encrypt_decrypt( desc) ,xor_encrypt_decrypt(time) ,xor_encrypt_decrypt(likecount ) ,xor_encrypt_decrypt(collectedcount) ,xor_encrypt_decrypt( commentcount ),
                                        xor_encrypt_decrypt(sharecount) ,xor_encrypt_decrypt(image)))
 
    conn.commit()
def InsertNoteHandleTocacheEncry( datas):
    global cursorsql
    # 定义插入单条数据的 SQL 语句
    insert_single_sql = '''INSERT INTO NodeHandleInfoEncry (noteID ,handleUserID , handleUserName, handleUserImage, handleType , handleTime ,mentionContent ,status,addtime,fans)
     VALUES (?,?,?,?,?,?,?,?,?,?)'''
    toinsert=[]
    for data in datas:
        toinsert.append((xor_encrypt_decrypt(data["篇"] ) ,xor_encrypt_decrypt(data["操作人ID"] ) ,xor_encrypt_decrypt(data["操作人昵称"]) ,xor_encrypt_decrypt(data["操作人头像"]) ,xor_encrypt_decrypt(data["操作类型"]
                       ) ,xor_encrypt_decrypt(data["操作时间"]) ,xor_encrypt_decrypt(data["评论内容"]) ,xor_encrypt_decrypt("1") ,xor_encrypt_decrypt(datetime.datetime.now().strftime("%Y_%m_%d_%H_%M_%S")),xor_encrypt_decrypt(data["关注"])))
    # 插入单条数据
    cursorsql.executemany(insert_single_sql, toinsert)
 
    conn.commit()
def InsertNoteInfoTocache(note_id ,xsec_token ,user_id , nickname="", title="", desc="" , time="" ,likecount=0 ,collectedcount=0, commentcount=0 ,sharecount=0,image=""):
    global cursorsql,InEncry,InNormal
    if(InNormal):
        # 定义插入单条数据的 SQL 语句
        insert_single_sql = '''INSERT INTO NodeTextInfo (note_id ,xsec_token ,user_id , nickname, title, desc , time ,likecount ,collectedcount, commentcount ,sharecount,image)
        VALUES (?, ?,?,?, ?,?,?, ?,?,?, ?,?)'''
        # 插入单条数据
        cursorsql.execute(insert_single_sql, (note_id ,xsec_token ,user_id , nickname, title, desc , time ,likecount ,collectedcount, commentcount ,sharecount,image))
    if(InEncry):
        InsertNoteInfoTocacheEncry(note_id ,xsec_token ,user_id , nickname, title, desc , time ,likecount ,collectedcount, commentcount ,sharecount,image)
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
def InsertNoteHandleTocache( datas):
    global cursorsql,InEncry,InNormal,noteToCalDetail,noteToCal,notehandletimeNo
    select_sql = "SELECT  ID, noteID,handleUserID,handleUserName,handleUserImage,handleType,handleTime,mentionContent,status,addtime FROM NodeHandleInfo" 
    cursorsql.execute(select_sql)
    # 获取所有查询结果
    dataNodeDZ1 = cursorsql.fetchall() 
     
    if(InNormal):  
        select_sql = "SELECT id,note_id,xsec_token,user_id,nickname,title,desc,time,likecount,collectedcount,commentcount,sharecount,image  FROM NodeTextInfo"
        # 执行查询语句
        cursorsql.execute(select_sql)
        # 获取所有查询结果
        NodeTexts = cursorsql.fetchall()
        # 定义插入单条数据的 SQL 语句
        insert_single_sql = '''INSERT INTO NodeHandleInfo (noteID ,handleUserID , handleUserName, handleUserImage, handleType , handleTime ,mentionContent ,status,addtime,fans)
        VALUES (?,?,?,?,?,?,?,?,?,?)'''
        toinsert=[]
        countDT={}#当天发布那篇的赞藏数，用来处理点的超过50的数据status置0
        for data in datas:
            status=1
            notehandletime=datetime.datetime.strptime(data["操作时间"], "%Y-%m-%d %H:%M:%S")
#-----------------------------------------------------------------------------处理新发的当天只收50赞50藏----------------------------------------------
            currentHandleDate=notehandletime.date()
            noteid=[id for id in NodeTexts if id[1]==data["篇"]][0]
            key=noteid#data["篇"]+data["操作类型"]
            if( key not in countDT): 
                countDT[key]=[0,0,0] 
            if(data["操作类型"]=="赞"):
                countDT[key][0]+=1
            elif(data["操作类型"]=="收藏"):
                countDT[key][1]+=1
            elif(data["操作类型"]=="评论"):
                countDT[key][2]+=1   
            datanode=[datac for datac in NodeTexts if datac[1]==data["篇"] ][0]        
            if((countDT[key][0]>32 or countDT[key][1]>32) and currentHandleDate in endtimes and  datetime.datetime.strptime(datanode[7], "%Y-%m-%d %H:%M:%S").date()==currentHandleDate):
                status=0            
#-------------------------------------------------------------------------------处理重复操作了的数据--------------------------------------------------
            if(len([dataC for dataC in toinsert  if dataC[0]==data["篇"] and dataC[1]==data["操作人ID"] and dataC[4]==data["操作类型"]])>0):
                status=0
                print(f'******{data["操作人昵称"]} 对篇{data["篇"]} 操作重复了{data["操作类型"]}')
#-----------------------------------------------------------------------------处理只要点赞或者收藏等不要全部的情况;处理某个时间后不再收赞或藏或评了--------------------------------------
            for i,ele in enumerate(noteToCal):
                if data["篇"]==ele:
                    notetimeNoList=notehandletimeNo[i].replace("\n","").replace("\r", "").split(";")
                    try:
                        if  data["操作类型"]=="赞":
                            if noteToCalDetail[i][0]!="1":
                                status=0
                            if(notehandletime>datetime.datetime.strptime(notetimeNoList[0], "%Y/%m/%d %H:%M:%S")):
                                status=0
                        elif data["操作类型"]=="收藏":
                            if noteToCalDetail[i][1]!="1":
                                status=0
                            if(notehandletime>datetime.datetime.strptime(notetimeNoList[1], "%Y/%m/%d %H:%M:%S")):
                                status=0
                        elif data["操作类型"]=="评论":
                            if noteToCalDetail[i][2]!="1":
                                status=0
                            if(notehandletime>datetime.datetime.strptime(notetimeNoList[2], "%Y/%m/%d %H:%M:%S")):
                                status=0
                    except Exception as ex:
                        print(ex)
                        traceback.print_exc()                        
                if(status==0):    
                    print(f'******{data["操作人昵称"]} 对篇{data["篇"]} 操作 {data["操作类型"]} 不和要求')
                    break
#-----------------------------------------------------------------------------处理没有在要处理的篇列表中的数据--------------------------------------
            if( data["篇"] not in  noteToCal ):
                status=0
            findold=[da for da in dataNodeDZ1 if (da[1]==data["篇"] and da[2]==data["操作人ID"] and da[5]==data["操作类型"] )]
            if (len(findold)>0):
                print(f'******你的账号 {data["操作人昵称"]} 对笔记《 {datanode[5]} 》的 {data["操作类型"]} 与过往重复，时间:{data["操作时间"]} ,{findold[0][6]}')
                status=0
            toinsert.append((data["篇"] , data["操作人ID"] ,data["操作人昵称"] ,data["操作人头像"],data["操作类型"],data["操作时间"],data["评论内容"],status,datetime.datetime.now().strftime("%Y_%m_%d_%H_%M_%S"),data["关注"]))
        # 插入单条数据
        cursorsql.executemany(insert_single_sql, toinsert)
        for key in countDT:
            print(f"{key[0]}  {key[5]} 有{countDT[key]}个")
    if(InEncry):
        InsertNoteHandleTocacheEncry(datas)
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
def xor_encrypt_decrypt(text, key=1123):
    encrypted_text = ""
    for char in str(text):
        encrypted_text += chr(ord(char) ^ key)
    return encrypted_text
if __name__ == '__main__':
    try:  
        #pyinstaller --onefile your_script.py

        global xhs_client
        global handleType,noteToCal,endtimes,InNormal,InEncry,noteToCalDetail,notehandletimeNo
        InEncry=False
        InNormal= not InEncry
        cookie = " "
        #--------------------------------------------------------------------------读配置的参数--------------------------------------------
        file_path='config\\config.csv'
        dataread=[]
        endtimes=[]
        if os.path.exists(file_path): 
            with open(file_path, mode='r', newline='', encoding='utf-8') as file:
                reader = csv.reader(file)
                dataread = list(reader)
                cookie = dataread[0][0]
                noteToCal=dataread[1]
                timetohandle=dataread[5]
                if(len(dataread)<5 or timetohandle[0]=="" or (timetohandle[0]!="" and datetime.date.today()< datetime.datetime.strptime(timetohandle[0], "%Y/%m/%d").date())):
                    endtimes.append(datetime.date.today()- datetime.timedelta(days=1)) #
                else:
                    endtimes=[datetime.datetime.strptime(datadate, "%Y/%m/%d").date() for datadate in timetohandle if datadate!=""]
                catchlike= int(dataread[4][0])
                catchMention=int (dataread[4][1])  
                noteToCalDetail=dataread[2]
                notehandletimeNo=dataread[3] 
        xhs_client = XhsClient(cookie, sign=sign)
        # 连接到 SQLite 数据库，如果数据库文件不存在则会创建一个新的数据库文件
        conn = sqlite3.connect('config\\WorkData.db')
        # 创建一个游标对象，用于执行 SQL 语句
        global cursorsql
        cursorsql = conn.cursor() 
        print(datetime.datetime.now())
        handleType={'faved/item':'收藏','liked/item':'赞',"comment/comment":'评论评论',"comment/item":'评论'}#"mention/item":'在笔记中@了你''liked/comment':赞了你的评论
        handleTypePrice={'faved/item':0.5,'liked/item':1,"comment/comment":0,"comment/item":0.5}
        fieldnames = ['操作ID','预览图','篇','篇title','作者','操作人ID','操作人昵称','操作人头像','操作类型',  '操作时间','评论内容','价格']
        data =GetInfoBySeq(catchlike,catchMention)
        if len(data)==0:
            print("没有获取到新的数据")
            exit(0) 
        sorted_list = sorted(data, key=lambda x: datetime.datetime.strptime(x["操作时间"], "%Y-%m-%d  %H:%M:%S"))
        InsertNoteHandleTocache(sorted_list) 
    except Exception as ex:
        print(ex)
        traceback.print_exc()
    finally:
        # 关闭游标
        cursorsql.close()
        # 关闭数据库连接
        conn.close()
