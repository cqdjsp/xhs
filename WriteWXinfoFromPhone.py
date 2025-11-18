import xlwings as xw
import uiautomator2 as u2
import time
import logging
import sqlite3 
import random
import sys
import re
from datetime import datetime
import os
#从手机中获取他们发的信息插入到数据库中
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class WechatVersion:
    def __init__(self, version):
        """初始化微信版本检查"""
        self.version = version 
        #默认版本为8.0.42
        self.User="com.tencent.mm:id/lpa"#用户 
        self.Content="com.tencent.mm:id/lp8"#用户发的信息
        self.Time="com.tencent.mm:id/lp_"#发信息时间
        self.ControlList="com.tencent.mm:id/lpg"#可滚动的信息控件
        self.ControlHole="com.tencent.mm:id/lp6"#用户，事件，信息所在的父控件
       

        if(version=="8.0.54"):
            self.User="com.tencent.mm:id/lpa"#用户 
            self.Content="com.tencent.mm:id/lp8"#用户发的信息
            self.Time="com.tencent.mm:id/lp_"#发信息时间
            self.ControlList="com.tencent.mm:id/lpg"#可滚动的信息控件
class WeChatDonation:
    def __init__(self, excel_path, password=None,startindex=0,versionWC=WechatVersion("8.0.42")):
        """初始化赞赏助手，加载Excel数据"""
        self.excel_path = excel_path
        self.password = password  # 支付密码，如需要  
        self.d = None  # uiautomator2设备对象
        self.startindex=startindex
        self.OBJWC=versionWC
 
    def connect_device(self):
        """连接Android设备"""
        try:
            self.d = u2.connect()  # 连接默认设备

            logger.info(f"已连接设备: {self.d.device_info}")
            return True
        except Exception as e:
            logger.error(f"连接设备失败: {str(e)}")
            return False
    
    def open_wechat(self):
        """打开微信应用"""
        try:
            self.d.app_start("com.tencent.mm")
            logger.info("正在打开微信...")
            time.sleep(3)  # 等待微信启动
            return True
        except Exception as e:
            logger.error(f"打开微信失败: {str(e)}")
            return False
    
def extract_and_convert_time(input_str):
    # 使用正则表达式提取日期时间部分（匹配类似2025_10_16_09_17_09的格式）
    match = re.search(r'(\d{4})_(\d{2})_(\d{2})_(\d{2})_(\d{2})_(\d{2})', input_str)
    
    if not match:
        raise ValueError("未在字符串中找到有效的时间格式")
    
    # 解析匹配到的时间部分
    year, month, day, hour, minute, second = map(int, match.groups())
    
    # 转换为datetime对象
    dt = datetime(year, month, day, hour, minute, second)
    
    # 格式化为目标格式（年.月.日 时:分:秒）
    return dt
def get_all_filenames(directory):
    # 检查目录是否存在
    if not os.path.exists(directory):
        raise FileNotFoundError(f"目录不存在: {directory}")
    
    if not os.path.isdir(directory):
        raise NotADirectoryError(f"{directory} 不是一个目录")
    
    # 获取目录下的所有文件和子目录
    all_entries = os.listdir(directory)
    
    # 筛选出文件（排除目录） 
    lastestPath=None
    lastest_time=datetime(2001, 1, 1,1,1,1)
    for entry in all_entries:
        entry_path = os.path.join(directory, entry)
        if os.path.isfile(entry_path):
            filenameTime=extract_and_convert_time(entry)
            if(filenameTime>lastest_time):#只处理未来的文件
                lastest_time=filenameTime
                lastestPath=entry_path   
    return lastestPath
if __name__ == "__main__":
 
    # Excel文件路径，确保文件存在且格式正确
    excel_path = get_all_filenames("E:/my/job/xhs/Result" ) #    excel_path = "E:/my/job/xhs/Result/结算(23-23)2025_06_24_09_43_20.xls"  
    startindex=2-2        #excel表格的行号-2
    versionWC=WechatVersion("8.0.42") #微信版本号
    d = u2.connect() # 连接多台设备需要指定设备序列号
    toinsertInfo=[]
    toinsertInfo1=[]
    cachesenders=[""]
    cachetimes=[""]
    cachecontents=[""]
    allin=False
    while(allin==False):
        allin=True
        controlHoles=d(resourceId=versionWC.ControlHole)
        for i in range(0,controlHoles.count):
            user33= controlHoles[i].child(resourceId=versionWC.User)
            time33= controlHoles[i].child(resourceId=versionWC.Time)
            content33=controlHoles[i].child(resourceId=versionWC.Content)
            contenth=""
            sender=""
            timeh=""
            if(user33.exists):
                sender=user33[0].get_text()
            if(time33.exists):
                timeh=time33[0].get_text()
            if(content33.exists):
                contenth=content33[0].get_text()
            find=False
            find1=False
            for valueO in toinsertInfo1:
                if(contenth!=""):
                    if( valueO[3]==contenth):
                        find1=True
                        break 

            if(find1==False and contenth!=""):
                toinsertInfo1.append((sender,"[聊天记录]" ,'text',contenth,timeh))
                allin=False




        #     for index,times in enumerate(cachetimes):
        #         if(timeh==times):
        #             if(cachesenders[index]==sender   ):
        #                 if((index<len(cachecontents) and cachecontents[index]==contenth)):
        #                     find=True
        #                     break
                

        #     if(find):
        #         continue
        #     else:
        #         allin=False
        #         if(sender!="" and cachesenders[-1]!=sender):
        #             cachesenders.append(sender)
        #         if(timeh!="" and cachetimes[-1]!=timeh):
        #             cachetimes.append(timeh )
        #         if(contenth!="" and cachecontents[-1]!=contenth):
        #             cachecontents.append(contenth)

        # # users= d(resourceId=versionWC.User)
        # # times= d(resourceId=versionWC.Time)
        # # contents= d(resourceId=versionWC.Content)
        # # for i in range(0,users.count):
        # #     sender=users[i].get_text()
        # #     if(sender not in cachesenders):
        # #         cachesenders.append(sender)
        # #         allin=False
        # # for i in range(0,times.count):
        # #     sender=times[i].get_text()
        # #     if(sender not in cachetimes):
        # #         cachetimes.append(sender)
        # #         allin=False
        # # for i in range(0,contents.count):
        # #     sender=contents[i].get_text()
        # #     if(sender not in cachecontents):
        # #         cachecontents.append(sender)
        # #         allin=False 
        d(resourceId=versionWC.ControlList).swipe("up",20)
        #
    for index,timeh in enumerate(cachetimes):
         toinsertInfo.append((cachesenders[index],"[聊天记录]" ,'text',cachecontents[index],cachetimes[index]))
    conn = sqlite3.connect('config\\WorkData.db')
    cursorsql = conn.cursor()
    insert_single_sql = '''INSERT INTO WXMSG (sender,myType,type,content,time)
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
    select_sql = "SELECT id,sender,myType,type,content,time FROM WXMSG WHERE" 
    cursorsql.execute(select_sql)
    # 获取所有查询结果
    dataNode2 = cursorsql.fetchall()  
    # 授予存储权限
    d.shell("pm grant com.github.uiautomator android.permission.WRITE_EXTERNAL_STORAGE")
    d.shell("pm grant com.github.uiautomator android.permission.READ_EXTERNAL_STORAGE") 
    print(d.info)
    # 创建支付助手实例
    donation = WeChatDonation(excel_path, password="705464",startindex=startindex,versionWC=versionWC)  # 替换为实际支付密码或留空
    try:
    # 执行支付
        donation.process_payments()
    except Exception as e:
        logger.error(f"支付过程中发生错误: {str(e)}")
        sys.exit(1)