# 主群处理
from wxautoMy.wxauto  import WeChat
import xlwings as xw
import sqlite3 
import datetime
import os
from PIL import Image
from pathlib import Path
#---------------------------------------------------------------------从excel中获取微信号插入pay图
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

def insertPayPicFromExcel():
    #从excel读取数据然后插入paycode的图片
    conn = sqlite3.connect('config\\WorkData.db')
    global cursorsql,sht,wb
    cursorsql = conn.cursor()
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False    # 关闭一些提示信息，可以加快运行速度。 默认为 True。
    app.screen_updating = True    # 更新显示工作表的内容。默认为 True。关闭它也可以提升运行速度。
    wb = app.books.open("E:\\my\\job\\xhs\\Result\\微信信息2025_03_07_15_04_52.xls")#xw.Book()#  
    sht = wb.sheets[0] 
    range_values = sht.range('A2:E80').value

    conn = sqlite3.connect('config\\WorkData.db')
    cursorsql = conn.cursor() 
    select_sql = "SELECT * FROM MarkWX" 
    cursorsql.execute(select_sql)
    # 获取所有查询结果
    dataNode2 = cursorsql.fetchall()  

    priceHeader=['用户名','用户ID','总额','支付码',"总额2","支付码"]    
    wb2 = xw.Book()# app.books.open('结算.xls') 
    sht2 = wb2.sheets[0] 
    # 将a1,a2,a3输入第一列，b1,b2,b3输入第二列
    sht2.range('A1') .value=priceHeader 
    i=2
    for row in range_values:
        id= row[1]
        sht2.range(f'A{i}') .value=list(row)
    # sht.range('B2') .options(transpose=True).value=list(handuserNicName.values())
    # sht.range('C2') .options(transpose=True).value=list(handuserPrice.values())
    # i=2 
        path=f'config\\zfcode\\{int(id)}.jpg'
        if(os.path.exists(path)):
            filePath = os.path.join(os.getcwd(),path )
            add_center(sht2, 'F'+str(i), filePath, width=350, height=350)
        i+=1
    wb2.save( "Result\\结算"+datetime.datetime.now().strftime("%Y_%m_%d_%H_%M_%S")+".xls" )
    wb2.close()
    wb.close()
    app.quit()

def paycodePicToCache():
    # 从支付的图片更新数据库id
    folder_path = Path('E:\\my\\job\\xhs\\config\\zfcode')

    # 查找文件夹内的所有文件
    files = [file.name.replace(".jpg","") for file in folder_path.iterdir() if file.is_file()]

    print("文件夹内的文件有：", files) 
    select_sql = "SELECT * FROM MarkWX WHERE PayCode<=0" 
    cursorsql.execute(select_sql)
    toupdateData=[]
    dataNode2 = cursorsql.fetchall()  
    for datanode in dataNode2:
        if(str(datanode[2]) in files):
            toupdateData.append((datanode[0],datanode[1],datanode[2],datanode[2]))
    for updateData in toupdateData: 
        print(f"更新了 {str(updateData[2])}的为{str(updateData[3])}")
        update_sql = "UPDATE MarkWX SET PayCode = ? WHERE MarkID = ?"
        cursorsql.execute(update_sql, (updateData[3], updateData[2]))
        # 提交事务，将更改保存到数据库
        conn.commit()

def downloadPayCodePic():
    #从主窗口下载图片
    wx = WeChat()
    select_sql = "SELECT * FROM MarkWX" 
    cursorsql.execute(select_sql)
    # 获取所有查询结果
    dataNode2 = cursorsql.fetchall()  
    msgs = wx.GetAllMessage(
        savepic   = True,   # 保存图片
        savefile  = False,   # 保存文件
        savevoice = False,    # 保存语音转文字内容
        saveVideo=False,
        saveZF=False,
        odata=dataNode2
    ) 
    # id=-1
    # savepath=""
    # if(odata!=None):
    #     for od in odata:
    #         if(od[1]==msg.sender):
    #             id=od[2]
    #             break
    #     if(id==-1):
    #         for od in odata:
    #             if((od[1] in msg.sender or msg.sender in od[1])):
    #                 id=od[2]
    #                 break
    # if(id!=-1):
    #     savepath=f"E:\\my\\job\\xhs\\config\\zfcode\\{id}.jpg"
    #     if(os.path.exists(savepath)):
    #         savepath=f"E:\\my\\job\\xhs\\config\\zfcode\\{id}__"+datetime.datetime.now().strftime("%Y_%m_%d_%H_%M_%S")+".jpg"

    #id=-1
    #savepath=""
    # name="快乐"
    # if(dataNode2!=None):
    #     for od in dataNode2:
    #         if(od[1]==name):
    #             id=od[2]
    #             break
    #     if(id==-1):
    #         for od in dataNode2:
    #             if((od[1] in name or name in od[1])):
    #                 id=od[2]
    #                 break
    # 获取当前聊天窗口消息
def updateWXtoXHS():
    #更新WXToXHSInfo的paycode，从markwx的paycode
    savepath=""
    select_sql = "SELECT * FROM WXToXHSInfo" 
    cursorsql.execute(select_sql) 
    dataNode = cursorsql.fetchall()  
    select_sql = "SELECT * FROM MarkWX" 
    cursorsql.execute(select_sql)
    toupdateData=[]
    dataNode2 = cursorsql.fetchall()  
    for dn in dataNode: 
        name=dn[1]
        id=-1
        if(dataNode2!=None):
            for od in dataNode2:
                if(od[1]==name):
                    id=od[3]
                    break
            if(id==-1):
                for od in dataNode2:
                    if((od[1] in name or name in od[1])):
                        id=od[3]
                        break 
        if(id!=-1):
            update_sql = "UPDATE WXToXHSInfo SET PayCode = ? WHERE id = ?"
            cursorsql.execute(update_sql, (id, dn[0]))
            # 提交事务，将更改保存到数据库
            conn.commit()               
def UpdateMarkFromLB():
    #从窗口中的群成员信息添加到数据库，先点击群成员的某一个人，要定住那个群成员窗口

    # update_sql = "UPDATE MarkWX SET PayCode = 0,MarkID=ID WHERE ID > 208"
    # cursorsql.execute(update_sql)
    # # 提交事务，将更改保存到数据库
    # conn.commit()  

    conn = sqlite3.connect('config\\WorkData.db')             
    global cursorsql,sht,wb
    cursorsql = conn.cursor() 
    uapi=wx.UiaAPI.ListControl(Name="聊天成员")
    uapic=uapi.GetChildren()

    select_sql = "SELECT * FROM MarkWX" 
    cursorsql.execute(select_sql)
    # 获取所有查询结果
    odata = cursorsql.fetchall()  
    savepath=""

    for uc in uapic:
        name=uc.Name 
        id=-1
        for od in odata:
            if(od[4]==name):
                id=od[0]
                break
        # if(id==-1):
        #     for od in odata:
        #         if((od[1] in name or name in od[1])):
        #             id=od[0]
        #             break
        if(id!=-1):
        #     update_sql = "UPDATE MarkWX SET OriginName = ? WHERE ID = ?"
        #     cursorsql.execute(update_sql, (name,id))
        #     # 提交事务，将更改保存到数据库
        #     conn.commit()  
        # else:
            MarkID=(odata[-1][0]+1 if cursorsql.lastrowid==0 else cursorsql.lastrowid) 
            print(f"更新了{name}，他MarkID是{MarkID}")
            insert_single_sql = '''INSERT INTO MarkWX (wxName ,MarkID ,PayCode,OriginName,AddTime)
            VALUES (?, ?,?,?,?)'''  
            cursorsql.execute(insert_single_sql, (name,MarkID,0,name,datetime.datetime.now().strftime("%Y_%m_%d_%H_%M_%S")))
            # 提交事务，将更改保存到数据库
            conn.commit()  
            
#------------------------------------------------------处理从聊天框获取支付码的图片


# priceZ=1
# priceC=0.5
# priceP=0.5     
# wx = WeChat()
# conn = sqlite3.connect('config\\WorkData.db')             
# global cursorsql,sht,wb
# cursorsql = conn.cursor()  
# downloadPayCodePic()
# paycodePicToCache() 

def remove_chars_around_colon(s):
    colon_index = s.find(':')
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

wx = WeChat()
tempMsg=wx.GetAllMessage(
        savepic   = False,   # 保存图片
        savefile  = False,   # 保存文件
        savevoice = False,    # 保存语音转文字内容
        saveVideo=False,
        saveZF=True,
        AreaText=("AreaText")
    ) 
# dataNodeDZ1Failed = [
#     [1, 2, 3, 4, 5],
#     [6, 7, 8, 9, 10],
#     [11, 12, 13, 14, 15]
# ]

# # 使用列表推导式提取每个子列表的第 4 个元素
# result = [data[3] for data in dataNodeDZ1Failed if data[2]>5]
# print(result)