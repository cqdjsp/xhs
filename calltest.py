# from xhs.core import XhsClient
# xhs_client = XhsClient(cookie="abRequestId=b1a8204b-f169-5ac9-a240-d7e76f92e284; xsecappid=xhs-pc-web; a1=192bd6bf1cfstrjudljds60zw3ua7ycqcd1hniisp50000115806; webId=aa08832c525b96208379fb35dcbb81eb; gid=yj2yJ8S4q0SjyjJDfKDiyf33SiM9FV7f1VfMyMUK8uEq7x280WvSAI888yy2Y8K820iSyWdi; unread={%22ub%22:%2267945d770000000018018490%22%2C%22ue%22:%226795a56e000000002900aa22%22%2C%22uc%22:24}; web_session=040069b73da8ec64049b58ca80354bb6cf315f; webBuild=4.57.0; loadts=1739950130583; websectiga=3fff3a6f9f07284b62c0f2ebf91a3b10193175c06e4f71492b60e056edcdebb2; sec_poison_id=59cc990d-908b-476d-86b6-f57ba5168334", # 用户 cookie
#                     ) # 自定义代理
# xhs_client.get_self_info2()
import datetime
import json
from time import sleep
import csv
import random
import re
from playwright.sync_api import sync_playwright
import datetime
from xhs.core import DataFetchError, XhsClient
from xhs import help
import os
import xlwings as xw
from PIL import Image
import sqlite3 
import traceback

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
 

def sign(uri, data=None, a1="", web_session=""):
    for _ in range(10):
        try:
            with sync_playwright() as playwright:
                stealth_js_path = "E:\\my\\job\\xhsTG\\public\\stealth.min.js"
                chromium = playwright.chromium

                # 如果一直失败可尝试设置成 False 让其打开浏览器，适当添加 sleep 可查看浏览器状态
                browser = chromium.launch(headless=True)

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
    # 定义查询数据的 SQL 语句
    select_sql = "SELECT * FROM NodeTextInfo"
    # 执行查询语句
    global cursorsql,noteToCal,endtimes
    cursorsql.execute(select_sql)
    # 获取所有查询结果
    dataNode = cursorsql.fetchall()
    data = [ ] 
    noteinfoTitle=""
    global handleType
    global xhs_client
    cursor=""
    while(catchlike>0): 
        note=  xhs_client.get_like_notifications(20,cursor)
        #note = xhs_client.get_note_by_id_from_html("67afecdf0000000028028c36","ABMuGHPzkrF3_R-x2Hv5gsctdOl93DbPpH4QcpptsADdg=")#,
        #print(json.dumps(note, indent=4))
        #print(help.get(note))
        findendtime=False#找到停止的时间也不再找了。退出while
        incurrectDay=False
        for noteInfo in note['message_list']:
            incurrectDay=False
            for endtime in endtimes: 
                if datetime.datetime.fromtimestamp(noteInfo['time']).date()!=endtime.date():
                    if datetime.datetime.fromtimestamp(noteInfo['time'])<endtime:
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
                    if(noteID not in  noteToCal):
                        continue
                    noteSec=noteInfo['item_info']['xsec_token'] if 'liked/item' ==noteInfo['type'] else noteInfo['item_info']["attach_item_info"]["xsec_token"] 
                    find=False
                    for notedatacache in dataNode:
                        if(notedatacache[1]==noteID):
                            noteinfoTitle=notedatacache[5]
                            find=True
                            break
                    if(find==False):
                        noteinfoSWeb=xhs_client.get_note_by_id(noteID,noteSec)
                        InsertNoteInfoTocache(noteID,noteSec,noteinfoSWeb["user"]["user_id"],noteinfoSWeb["user"]["nickname"],noteinfoSWeb["title"],noteinfoSWeb["desc"],datetime.datetime.fromtimestamp(noteinfoSWeb['time']/1000).strftime("%Y-%m-%d %H:%M:%S"),
                                              int(noteinfoSWeb["interact_info"]["liked_count"]),int(noteinfoSWeb["interact_info"]["collected_count"]),int(noteinfoSWeb["interact_info"]["comment_count"]),int(noteinfoSWeb["interact_info"]["share_count"]),noteInfo['item_info']['image'])
                        noteinfoTitle=noteinfoSWeb["title"]
                        cursorsql.execute(select_sql)
                        # 获取所有查询结果
                        dataNode = cursorsql.fetchall()
                    data.append({"预览图":noteInfo['item_info']['image'],"篇":noteID,"篇title":noteinfoTitle,
                                "作者":noteInfo['item_info']['user_info']["userid"] if  'user_info' in noteInfo['item_info'] else "",
                                '操作人ID':noteInfo["user_info"]['userid'], '操作人昵称':noteInfo["user_info"]['nickname'],'操作人头像':noteInfo["user_info"]['image'],
                                '操作类型':handleType[noteInfo['type']],'操作时间':datetime.datetime.fromtimestamp(noteInfo['time']).strftime("%Y-%m-%d %H:%M:%S"),'价格':handleTypePrice[noteInfo['type']],
                                '评论内容': "",'操作ID': str(noteInfo['time']) +'_'+noteInfo['id']
                            }  )
        if(findendtime): break    
        catchlike-=20
        cursor=note['strCursor']  
        if(note["has_more"]==False):break  
        sleep(random.randint(1, 3))
    cursor=""
    while(catchMention>0):
        mentionNote=  xhs_client.get_mention_notifications(20,cursor)
        findendtime=False#找到停止的时间也不再找了。退出while
        for noteInfo in mentionNote['message_list'] :
            incurrectDay=False
            for endtime in endtimes: 
                if datetime.datetime.fromtimestamp(noteInfo['time']).date()!=endtime.date():
                    if datetime.datetime.fromtimestamp(noteInfo['time'])<endtime:
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
                if( noteID not in  noteToCal ):
                    continue
              
                noteSec=noteInfo['item_info']['xsec_token'] if 'xsec_token' in noteInfo['item_info'] else ""
                find=False
                for notedatacache in dataNode:
                    if(notedatacache[1]==noteID):
                        noteinfoTitle=notedatacache[5]
                        find=True
                        break
                if(find==False):
                    noteinfoSWeb=xhs_client.get_note_by_id(noteID,noteSec)
                    InsertNoteInfoTocache(noteID,noteSec,noteinfoSWeb["user"]["user_id"],noteinfoSWeb["user"]["nickname"],noteinfoSWeb["title"],noteinfoSWeb["desc"],datetime.datetime.fromtimestamp(noteinfoSWeb['time']/1000).strftime("%Y-%m-%d %H:%M:%S"),
                                            int(noteinfoSWeb["interact_info"]["liked_count"]),int(noteinfoSWeb["interact_info"]["collected_count"]),int(noteinfoSWeb["interact_info"]["comment_count"]),int(noteinfoSWeb["interact_info"]["share_count"]),noteInfo['item_info']['image'])
                    noteinfoTitle=noteinfoSWeb["title"]
                    cursorsql.execute(select_sql)
                    # 获取所有查询结果
                    dataNode = cursorsql.fetchall()
                data.append({"预览图":noteInfo['item_info']['image'],"篇":noteID,"篇title":noteinfoTitle,
                            "作者":noteInfo['item_info']['user_info']["userid"] if  'user_info' in noteInfo['item_info'] else "",
                            '操作人ID':noteInfo["user_info"]['userid'], '操作人昵称':noteInfo["user_info"]['nickname'],'操作人头像':noteInfo["user_info"]['image'],
                            '操作类型':handleType[noteInfo['type']],'操作时间':datetime.datetime.fromtimestamp(noteInfo['time']).strftime("%Y-%m-%d %H:%M:%S"),'价格':handleTypePrice[noteInfo['type']],
                            '评论内容': noteInfo['comment_info']["content"],'操作ID': noteInfo['comment_info']["id"]
                        }  )
        if(findendtime): break 
        catchMention-=20
        cursor=mentionNote['strCursor']  
        if(mentionNote["has_more"]==False):break  
        sleep(random.randint(1, 3))
    return data
def InsertNoteInfoTocache(note_id ,xsec_token ,user_id , nickname="", title="", desc="" , time="" ,likecount=0 ,collectedcount=0, commentcount=0 ,sharecount=0,image=""):
    global cursorsql
    # 定义插入单条数据的 SQL 语句
    insert_single_sql = '''INSERT INTO NodeTextInfo (note_id ,xsec_token ,user_id , nickname, title, desc , time ,likecount ,collectedcount, commentcount ,sharecount,image)
     VALUES (?, ?,?,?, ?,?,?, ?,?,?, ?,?)'''
    # 插入单条数据
    cursorsql.execute(insert_single_sql, (note_id ,xsec_token ,user_id , nickname, title, desc , time ,likecount ,collectedcount, commentcount ,sharecount,image))

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
    global cursorsql
    # 定义插入单条数据的 SQL 语句
    insert_single_sql = '''INSERT INTO NodeHandleInfo (noteID ,handleUserID , handleUserName, handleUserImage, handleType , handleTime ,mentionContent ,status,addtime)
     VALUES (?,?,?,?,?,?,?,?,?)'''
    toinsert=[]
    for data in datas:
        toinsert.append((data["篇"] , data["操作人ID"] ,data["操作人昵称"] ,data["操作人头像"],data["操作类型"],data["操作时间"],data["评论内容"],1,datetime.datetime.now().strftime("%Y_%m_%d_%H_%M_%S")))
    # 插入单条数据
    cursorsql.executemany(insert_single_sql, toinsert)

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
if __name__ == '__main__':
    try:  
        global xhs_client
        global handleType,noteToCal,endtimes
        cookie = "abRequestId=b1a8204b-f169-5ac9-a240-d7e76f92e284; a1=192bd6bf1cfstrjudljds60zw3ua7ycqcd1hniisp50000115806; webId=aa08832c525b96208379fb35dcbb81eb; gid=yj2yJ8S4q0SjyjJDfKDiyf33SiM9FV7f1VfMyMUK8uEq7x280WvSAI888yy2Y8K820iSyWdi; x-user-id-creator.xiaohongshu.com=6649eba4000000000d0254cd; customerClientId=244158184254652; access-token-creator.xiaohongshu.com=customer.creator.AT-68c517473327474294336871falud8fmecdtl0kl; galaxy_creator_session_id=vkGmDW3N11pOMMAOEiFNwrTTAT3diXO8XfED; galaxy.creator.beaker.session.id=1740019646207039144021; xsecappid=xhs-pc-web; webBuild=4.60.1; web_session=04006979478696029c2ce88fed354bd9723a03; unread={%22ub%22:%2267d37e04000000001c00d16c%22%2C%22ue%22:%2267d8074a000000000602ae7b%22%2C%22uc%22:16}; websectiga=984412fef754c018e472127b8effd174be8a5d51061c991aadd200c69a2801d6; sec_poison_id=598e0daa-8814-40f0-b0e2-cd8f745be798; loadts=1742269363475"
        catchlike=2000#获取100个赞藏数据
        catchMention=300#获取100个评论数据
        noteToCal=["","67d930db000000000d016ae5"]#,"67d546e200000000060284cb"
        endtimes=[datetime.date(2025, 3, 21),datetime.date(2025, 3, 22)]
        #file_path="Result\\"+datetime.datetime.now().strftime("%Y_%m_%d_%H_%M_%S")+"Detail.csv"
        xhs_client = XhsClient(cookie, sign=sign)
        # 连接到 SQLite 数据库，如果数据库文件不存在则会创建一个新的数据库文件
        conn = sqlite3.connect('config\\WorkData.db')
        # 创建一个游标对象，用于执行 SQL 语句
        global cursorsql
        cursorsql = conn.cursor()
        # 定义创建表的 SQL 语句
        #InsertNoteInfoTocache('6279927a0000000001028487','LBumpsGOchO7EzgOzA56KRfn32nNT9EvRcgwLfIPjFGYs=','5f6303c5000000000101ebb4','姜可可艾','有一种童年的向往，就是宫崎骏的夏天',"",'2022-05-10 06:15:22',2,1,2,4)
        create_table_sql = '''
        CREATE TABLE IF NOT EXISTS NodeTextInfo (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            note_id TEXT NOT NULL,
            xsec_token TEXT NOT NULL,
            user_id TEXT,
            nickname TEXT,
            title TEXT ,
            desc TEXT,
            time TEXT,
            likecount INTEGER,
            collectedcount INTEGER,
            commentcount INTEGER,
            sharecount INTEGER
        )
        '''
        # 执行 SQL 语句创建表
        cursorsql.execute(create_table_sql)
        # 提交事务，将更改保存到数据库
        conn.commit() 


        print(datetime.datetime.now())
        handleType={'faved/item':'收藏','liked/item':'赞',"comment/comment":'评论评论',"comment/item":'评论'}#"mention/item":'在笔记中@了你''liked/comment':赞了你的评论
        handleTypePrice={'faved/item':0.5,'liked/item':1,"comment/comment":0,"comment/item":0.5}
        fieldnames = ['操作ID','预览图','篇','篇title','作者','操作人ID','操作人昵称','操作人头像','操作类型',  '操作时间','评论内容','价格']
        data =GetInfoBySeq(catchlike,catchMention)
        InsertNoteHandleTocache(data)

        #dataread=[]
        # if os.path.exists(file_path): 
        #     with open(file_path, mode='r', newline='', encoding='utf-8') as file:
        #         reader = csv.DictReader(file)
        #         dataread = list(reader)
        #     if(len(dataread)>0):  
        #         with open(file_path, mode='a', newline='', encoding='utf-8') as file:
        #             writer = csv.DictWriter(file, fieldnames=fieldnames) 
        #             for toaddData in data:
        #                 existData=False
        #                 for datareadInfo in dataread:
        #                     if('操作ID' in datareadInfo and toaddData['操作ID']  ==datareadInfo['操作ID']):
        #                         existData=True
        #                         break
        #                 if(existData==False):
        #                     writer.writerow(toaddData)
        #     else: 
        #         with open(file_path, mode='w', newline='', encoding='utf-8') as file:
        #             writer = csv.DictWriter(file, fieldnames=fieldnames)
        #             writer.writeheader()
        #             writer.writerows(data)                        
        # else:
        #     with open(file_path, mode='x', newline='', encoding='utf-8') as file:
        #         writer = csv.DictWriter(file, fieldnames=fieldnames)
        #         writer.writeheader()
        #         writer.writerows(data)   

        # with open(file_path, mode='r', newline='', encoding='utf-8') as file:
        #     reader = csv.DictReader(file)
        #     dataread = list(reader)
        # handuserPrice={}
        # handuserNicName={}
        # handuserPriceToAdd=[]
        # datareadCode=[]
        # paycode={}
        # with open('config//IDToPayCode.csv', mode='r', newline='', encoding='utf-8') as file:
        #     reader = csv.DictReader(file)
        #     datareadCode = list(reader)
        # if(len(dataread)>0):
        #     for dataInfo in dataread: 
        #         paycodeSet=""
        #         for datas in datareadCode:
        #             if(datas["ID"]==dataInfo["操作人ID"]):
        #                 paycodeSet=datas["支付码ID"]
        #                 break
        #         paycode[dataInfo["操作人ID"] ]=paycodeSet
        #         if( dataInfo["操作人ID"]  in handuserPrice):
        #             handuserPrice[dataInfo["操作人ID"] ]+=float(dataInfo['价格'])
        #         else:
        #             handuserPrice[dataInfo["操作人ID"] ]=float(dataInfo['价格'])
        #             handuserNicName[dataInfo["操作人ID"] ]=dataInfo['操作人昵称']
                    
        #     for handinfo in handuserPrice:
        #         paycode[dataInfo["操作人ID"] ]=""
        #         for datas in datareadCode:
        #             if(datas["ID"]==handinfo):
        #                 paycode[dataInfo["操作人ID"] ]=datas["支付码ID"]
        #         handuserPriceToAdd.append({"用户ID":handinfo,'用户名':handuserNicName[handinfo],"总额":handuserPrice[handinfo],"支付码":paycode[dataInfo["操作人ID"] ]})
        #     priceHeader=['用户ID','用户名','总额','支付码']
        #     # with open('output.csv', mode='a', newline='', encoding='utf-8') as file:
        #     #     writer = csv.DictWriter(file,priceHeader) 
        #     #     writer.writeheader()
        #     #     writer.writerows(handuserPriceToAdd)
    
        #     app = xw.App(visible=True, add_book=False)
        #     app.display_alerts = False    # 关闭一些提示信息，可以加快运行速度。 默认为 True。
        #     app.screen_updating = True    # 更新显示工作表的内容。默认为 True。关闭它也可以提升运行速度。
        #     wb = xw.Book()# app.books.open('结算.xls') 
        #     sht = wb.sheets[0] 
        #     # 将a1,a2,a3输入第一列，b1,b2,b3输入第二列
        #     sht.range('A1') .value=priceHeader 
        #     sht.range('A2') .options(transpose=True).value=list(handuserPrice.keys())
        #     sht.range('B2') .options(transpose=True).value=list(handuserNicName.values())
        #     sht.range('C2') .options(transpose=True).value=list(handuserPrice.values())
        #     i=2
        #     for paycodekey in paycode:
        #         if(paycode[paycodekey]!=""):
        #             filePath = os.path.join(os.getcwd(), f'config\\zfcode\\{paycode[paycodekey]}.jpg')
        #             add_center(sht, 'D'+str(i), filePath, width=150, height=150)
        #         i+=1
        #     wb.save( "Result\\结算"+datetime.datetime.now().strftime("%Y_%m_%d_%H_%M_%S")+".xls" )
        #     wb.close()
            #wb.app.quit()
    except Exception as ex:
        print(ex)
        traceback.print_exc()
    finally:
        # 关闭游标
        cursorsql.close()
        # 关闭数据库连接
        conn.close()