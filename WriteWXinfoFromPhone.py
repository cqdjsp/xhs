import uiautomator2 as u2 
import logging
import sqlite3 
import datetime
import re
import time
from datetime import datetime 
#从手机中获取他们发的信息插入到数据库中
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)
from datetime import datetime, timedelta
import re

from datetime import datetime, timedelta
import re
 
def process_time(input_time: str) -> str:
    """
    处理输入时间：
    1. 纯时间格式（HH:MM:SS）→ 补充前一天日期（当前年份+前一天月日），格式 YYYY/MM/DD HH:MM:SS
    2. 无年份日期+时间（MM月DD日 HH:MM:SS）→ 补充当前年份，格式 YYYY/MM/DD HH:MM:SS
    3. 带年份日期+时间（YYYY年MM月DD日 HH:MM:SS）→ 直接转换为 YYYY/MM/DD HH:MM:SS
    :param input_time: 输入时间字符串（支持三种格式）
    :return: 处理后的时间字符串（格式："YYYY/MM/DD HH:MM:SS"）
    :raises ValueError: 输入格式不支持时抛出异常
    """
    # 定义三种格式的正则表达式
    # 1. 纯时间：HH:MM:SS（如 00:12:30）
    time_only_pattern = r'^(\d{1,2}):(\d{2}):(\d{2})$'
    # 2. 无年份日期+时间：MM月DD日 HH:MM:SS（如 12月30日 00:12:30）
    date_time_no_year_pattern = r'^(\d{1,2})月(\d{1,2})日\s+(\d{1,2}):(\d{2}):(\d{2})$'
    # 3. 带年份日期+时间：YYYY年MM月DD日 HH:MM:SS（如 2025年12月30日 00:12:30）
    date_time_with_year_pattern = r'^(\d{4})年(\d{1,2})月(\d{1,2})日\s+(\d{1,2}):(\d{2}):(\d{2})$'

    # 去除首尾空格，统一处理
    input_clean = input_time.strip()
    current_year = datetime.now().year

    # 匹配1：纯时间格式（补充前一天日期）
    time_match = re.match(time_only_pattern, input_clean)
    if time_match:
        yesterday = datetime.now().date() - timedelta(days=1)
        date_str = yesterday.strftime(f"{current_year}-%m-%d")
        # 补零确保时间部分为两位数（如 1:2:3 → 01:02:03）
        h, m, s = time_match.groups()
        time_part = f"{h.zfill(2)}:{m.zfill(2)}:{s.zfill(2)}"
        return f"{date_str} {time_part}"

    # 匹配2：带年份的日期+时间格式（直接转换）
    year_date_time_match = re.match(date_time_with_year_pattern, input_clean)
    if year_date_time_match:
        year, month, day, h, m, s = year_date_time_match.groups()
        # 补零确保月、日、时为两位数
        year = int(year)
        month = int(month)
        day = int(day)
        time_part = f"{h.zfill(2)}:{m.zfill(2)}:{s.zfill(2)}"
        date_str = f"{year}-{month:02d}-{day:02d}"
        return f"{date_str} {time_part}"

    # 匹配3：无年份的日期+时间格式（补充当前年份）
    no_year_date_time_match = re.match(date_time_no_year_pattern, input_clean)
    if no_year_date_time_match:
        month, day, h, m, s = no_year_date_time_match.groups()
        month = int(month)
        day = int(day)
        time_part = f"{h.zfill(2)}:{m.zfill(2)}:{s.zfill(2)}"
        date_str = f"{current_year}-{month:02d}-{day:02d}"
        return f"{date_str} {time_part}"

    # 不匹配任何格式，抛出异常
    raise ValueError(
        f"不支持的时间格式：{input_time}\n"
        f"支持格式：\n"
        f"1. 纯时间（如：11:39:58）\n"
        f"2. 无年份日期+时间（如：11月11日 2:49:58）\n"
        f"3. 带年份日期+时间（如：2025年12月30日 00:12:30）"
    )
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
        self.ZFContent="com.tencent.mm:id/cu2"#转发的内容

        if(version=="8.0.54"):
            self.User="com.tencent.mm:id/lpa"#用户 
            self.Content="com.tencent.mm:id/lp8"#用户发的信息
            self.Time="com.tencent.mm:id/lp_"#发信息时间
            self.ControlList="com.tencent.mm:id/lpg"#可滚动的信息控件
            self.ZFContent="com.tencent.mm:id/cu2"#转发的内容
  
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
if __name__ == "__main__":
  
    versionWC=WechatVersion("8.0.42") #微信版本号
    d = u2.connect() # 连接多台设备需要指定设备序列号 
    toinsertInfo1=[]
    cachesenders=[""]
    cachetimes=[""]
    cachecontents=[""]
    allin=False
    findtop=False
    tempValue=[]
    #----------------------找到聊天记录的最顶部---------------------------
    while(findtop==False):
        d(resourceId=versionWC.ControlList).swipe("down",10) 
        controlHoles=d(resourceId=versionWC.ControlHole)
        user33= controlHoles[0].child(resourceId=versionWC.User)[0].get_text()
        time33= controlHoles[0].child(resourceId=versionWC.Time)[0].get_text()
        content33=controlHoles[0].child(resourceId=versionWC.Content)[0].get_text()
        if(len(tempValue)==0):
            tempValue.append((user33,time33,content33))
        else:
            if(tempValue[0][0]==user33 and tempValue[0][1]==time33 and tempValue[0][2]==content33):
                tempValue.append((user33,time33,content33))
            else:
                tempValue.clear()
                tempValue.append((user33,time33,content33))
        if(len(tempValue)==4):
            findtop=True
    #---------------------
    CFData=[]
    while(allin==False):
        allin=True
        time.sleep(3)
        controlHoles=d(resourceId=versionWC.ControlHole)
        for i in range(0,controlHoles.count):
            user33= controlHoles[i].child(resourceId=versionWC.User)
            time33= controlHoles[i].child(resourceId=versionWC.Time)
            content33=controlHoles[i].child(resourceId=versionWC.Content)
            contenth=""
            sender=""
            timeh=""
            parentBounds=controlHoles[i].bounds()
            if(user33.exists):
                bou=user33[0].bounds()
                if(bou[1]>=parentBounds[1] and bou[3]<=parentBounds[3]):
                    sender=user33[0].get_text()
            if(time33.exists   ):
                bou=user33[0].bounds()
                if(bou[1]>=parentBounds[1] and bou[3]<=parentBounds[3]): 
                    timeh=process_time(time33[0].get_text())
            if(content33.exists):
                bou=content33[0].bounds()
                if(bou[1]>=parentBounds[1] and bou[3]<=parentBounds[3]): 
                    contenth=content33[0].get_text()
                    if(contenth==""):
                        content33=controlHoles[i].child(resourceId=versionWC.ZFContent)
                        if(content33.exists):
                            bou=content33[0].bounds()
                            if(bou[1]>=parentBounds[1] and bou[3]<=parentBounds[3]): 
                                contenth=content33[0].get_text().split(sender)[1]+"@姜可艾 没有结算完"
                                contenth=contenth.replace(":","")
            find=False
            find1=False
            for valueO in toinsertInfo1:
                if(contenth!=""):
                    if( valueO[3]==contenth):
                        if(valueO[3]==contenth and timeh!='' and valueO[4]!=timeh):
                            CFData.append((sender,timeh,contenth))
                            print(f"找到重复的数据 {timeh},{sender},{contenth}")
                        find1=True
                        break 

            if(find1==False and contenth!=""):
                toinsertInfo1.append((sender,"[聊天]" ,'text',contenth,timeh,datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                allin=False
 
        d(resourceId=versionWC.ControlList).swipe("up",20) 
 
    conn = sqlite3.connect('config\\WorkData.db')
    cursorsql = conn.cursor()
    insert_single_sql = '''INSERT INTO WXMSG (sender,myType,type,content,time,HandleDate)
     VALUES (?, ?,?,?,?,?)'''      
    cursorsql.executemany(insert_single_sql, toinsertInfo1)
    # 提交事务，将更改保存到数据库
    conn.commit()   
    print(d.info)
 