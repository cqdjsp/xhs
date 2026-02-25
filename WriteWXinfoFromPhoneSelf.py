#从做数据的群里直接获取，这会导致没有时间信息，时间都以开始日期代替
import uiautomator2 as u2 
import logging
import sqlite3 
import datetime
import re
import os
import csv
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
    1. 纯时间格式（HH:MM）→ 补充秒数为00，再补充前一天日期，格式 YYYY-MM-DD HH:MM:SS
    2. 纯时间格式（HH:MM:SS）→ 补充前一天日期，格式 YYYY-MM-DD HH:MM:SS
    3. 无年份日期+时间（MM月DD日 HH:MM:SS）→ 补充当前年份，格式 YYYY-MM-DD HH:MM:SS
    4. 带年份日期+时间（YYYY年MM月DD日 HH:MM:SS）→ 直接转换为 YYYY-MM-DD HH:MM:SS
    :param input_time: 输入时间字符串（支持四种格式）
    :return: 处理后的时间字符串（格式："YYYY-MM-DD HH:MM:SS"）
    :raises ValueError: 输入格式不支持时抛出异常
    """
    # 定义四种格式的正则表达式
    # 1. 纯时间：HH:MM（如 03:26）
    time_hhmm_pattern = r'^(\d{1,2}):(\d{2})$'
    # 2. 纯时间：HH:MM:SS（如 00:12:30）
    time_only_pattern = r'^(\d{1,2}):(\d{2}):(\d{2})$'
    # 3. 无年份日期+时间：MM月DD日 HH:MM（新增，如 1月23日 08:29）
    date_time_no_year_hhmm_pattern = r'^(\d{1,2})月(\d{1,2})日\s+(\d{1,2}):(\d{2})$'
    # 3. 无年份日期+时间：MM月DD日 HH:MM:SS（如 12月30日 00:12:30）
    date_time_no_year_pattern = r'^(\d{1,2})月(\d{1,2})日\s+(\d{1,2}):(\d{2}):(\d{2})$'
    # 4. 带年份日期+时间：YYYY年MM月DD日 HH:MM:SS（如 2025年12月30日 00:12:30）
    date_time_with_year_pattern = r'^(\d{4})年(\d{1,2})月(\d{1,2})日\s+(\d{1,2}):(\d{2}):(\d{2})$'

    # 去除首尾空格，统一处理
    input_clean = input_time.strip()
    current_year = datetime.now().year
    # 匹配1：纯时间格式 HH:MM（补充秒数00和前一天日期）
    time_hhmm_match = re.match(time_hhmm_pattern, input_clean)
    if time_hhmm_match:
        yesterday = datetime.now().date() - timedelta(days=1)
        date_str = yesterday.strftime(f"{current_year}-%m-%d")  # 保留-分隔符
        # 补零确保时间部分为两位数，补充秒数00
        h, m = time_hhmm_match.groups()
        time_part = f"{h.zfill(2)}:{m.zfill(2)}:00"
        return f"{date_str} {time_part}"
    # 匹配2：纯时间格式（补充前一天日期）
    time_match = re.match(time_only_pattern, input_clean)
    if time_match:
        yesterday = datetime.now().date() - timedelta(days=1)
        date_str = yesterday.strftime(f"{current_year}-%m-%d")
        # 补零确保时间部分为两位数（如 1:2:3 → 01:02:03）
        h, m, s = time_match.groups()
        time_part = f"{h.zfill(2)}:{m.zfill(2)}:{s.zfill(2)}"
        return f"{date_str} {time_part}"
    # 匹配3：无年份日期+时间 HH:MM（新增，补充秒数00和当前年份）
    date_time_no_year_hhmm_match = re.match(date_time_no_year_hhmm_pattern, input_clean)
    if date_time_no_year_hhmm_match:
        month, day, h, m = date_time_no_year_hhmm_match.groups()
        month = int(month)
        day = int(day)
        # 补零+补充秒数00
        time_part = f"{h.zfill(2)}:{m.zfill(2)}:00"
        date_str = f"{current_year}-{month:02d}-{day:02d}"
        return f"{date_str} {time_part}"
    # 匹配3：带年份的日期+时间格式（直接转换）
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

    # 匹配4：无年份的日期+时间格式（补充当前年份）
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
        #默认版本为8.0.60
        self.User="com.tencent.mm:id/brc"#用户 
        self.Content="com.tencent.mm:id/bkl"#用户发的信息
        self.Time="com.tencent.mm:id/br1"#发信息时间
        self.ControlList="com.tencent.mm:id/bp0"#可滚动的信息控件
        self.ControlHole="com.tencent.mm:id/bkj"#用户，事件，信息所在的父控件
        self.ZFContent="com.tencent.mm:id/bj2"#转发的内容

        if(version=="8.0.43"):
            self.User="com.tencent.mm:id/lpa"#用户 
            self.Content="com.tencent.mm:id/lp8"#用户发的信息
            self.Time="com.tencent.mm:id/lp_"#发信息时间
            self.ControlList="com.tencent.mm:id/lpg"#可滚动的信息控件
            self.ZFContent="com.tencent.mm:id/cu2"#转发的内容

def parse_wechat_time(time_str):
    """
    解析微信聊天记录的系统时间字符串，支持以下格式：
    1. 具体日期：2025年4月25日 5:48
    2. 星期+时间：星期一 9:40、周一16:12
    3. 昨天+时间：昨天 9:40、昨天 晚上6:20、昨天 下午4:14
    4. 仅时间：9:40、16:12
    5. 带时段时间：上午8:25、下午3:40、晚上10:15、凌晨12:30、中午12:10
    """
    # 预处理：去除首尾空格，统一中文星期（周一→星期一）
    time_str = time_str.strip()
    # 映射简写星期到完整星期
    week_map = {
        '周一': '星期一', '周二': '星期二', '周三': '星期三',
        '周四': '星期四', '周五': '星期五', '周六': '星期六',
        '周日': '星期天', '周日': '星期天'
    }
    for short, full in week_map.items():
        time_str = time_str.replace(short, full)

    # 获取当前日期和时间
    now = datetime.now()
    today = now.date()
    year = today.year
    month = today.month
    day = today.day

    # 定义正则模式（按匹配优先级排序，复合格式优先）
    patterns = [
        # 模式1：具体日期（如：2025年4月25日 5:48）
        r'^(\d{4})年(\d{1,2})月(\d{1,2})日\s*(\d{1,2}):(\d{2})$',
        
        # 模式2：昨天+时段+时间（如：昨天 晚上6:20、昨天 下午4:14）
        r'^昨天\s*(上午|下午|晚上|凌晨|中午)\s*(\d{1,2}):(\d{2})$',
        
        # 模式3：昨天+纯时间（如：昨天 9:40、昨天 16:12）
        r'^昨天\s*(\d{1,2}):(\d{2})$',
        
        # 模式4：星期+时段+时间（如：星期一 晚上8:20、星期一 中午12:10）
        r'^(星期一|星期二|星期三|星期四|星期五|星期六|星期天)\s*(上午|下午|晚上|凌晨|中午)\s*(\d{1,2}):(\d{2})$',
        
        # 模式5：星期+纯时间（如：星期一 9:40、星期一 16:12）
        r'^(星期一|星期二|星期三|星期四|星期五|星期六|星期天)\s*(\d{1,2}):(\d{2})$',
        
        # 模式6：带时段的纯时间（如：上午8:25、中午12:10、晚上10:15）
        r'^(上午|下午|晚上|凌晨|中午)\s*(\d{1,2}):(\d{2})$',
        
        # 模式7：纯时间（如：9:40、16:12）
        r'^(\d{1,2}):(\d{2})$'
    ]

    # 解析时段到24小时制的辅助函数
    def period_to_24h(period, hour):
        hour = int(hour)
        if period == '下午' and hour != 12:
            hour += 12
        elif period == '晚上' and hour != 12:
            hour += 12
        elif period == '凌晨' and hour == 12:
            hour = 0
        elif period == '中午' and hour == 12:
            hour = 12  # 中午12点保持12点
        # 确保小时数在0-23范围内
        return hour % 24

    # 遍历匹配模式
    for idx, pattern in enumerate(patterns):
        match = re.match(pattern, time_str)
        if not match:
            continue

        # 模式1：具体日期
        if idx == 0:
            y, m, d, h, mi = match.groups()
            return datetime(int(y), int(m), int(d), int(h), int(mi))
        
        # 模式2：昨天+时段+时间
        elif idx == 1:
            period, h, mi = match.groups()
            h_24 = period_to_24h(period, h)
            yesterday = today - timedelta(days=1)
            return datetime(yesterday.year, yesterday.month, yesterday.day, h_24, int(mi))
        
        # 模式3：昨天+纯时间
        elif idx == 2:
            h, mi = match.groups()
            yesterday = today - timedelta(days=1)
            return datetime(yesterday.year, yesterday.month, yesterday.day, int(h), int(mi))
        
        # 模式4：星期+时段+时间
        elif idx == 3:
            week_day, period, h, mi = match.groups()
            h_24 = period_to_24h(period, h)
            # 计算目标星期的日期（微信显示的是「过去的同星期」）
            current_weekday = now.weekday()  # 0=周一，6=周日
            target_weekday = ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期天'].index(week_day)
            delta_days = target_weekday - current_weekday
            if delta_days > 0:  # 目标星期在本周未来，取上周的同星期
                delta_days -= 7
            target_date = today + timedelta(days=delta_days)
            return datetime(target_date.year, target_date.month, target_date.day, h_24, int(mi))
        
        # 模式5：星期+纯时间
        elif idx == 4:
            week_day, h, mi = match.groups()
            current_weekday = now.weekday()
            target_weekday = ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期天'].index(week_day)
            delta_days = target_weekday - current_weekday
            if delta_days > 0:
                delta_days -= 7
            target_date = today + timedelta(days=delta_days)
            return datetime(target_date.year, target_date.month, target_date.day, int(h), int(mi))
        
        # 模式6：带时段的纯时间
        elif idx == 5:
            period, h, mi = match.groups()
            h_24 = period_to_24h(period, h)
            return datetime(year, month, day, h_24, int(mi))
        
        # 模式7：纯时间
        elif idx == 6:
            h, mi = match.groups()
            return datetime(year, month, day, int(h), int(mi))

    # 匹配失败抛出异常
    raise ValueError(f"无法解析的微信时间格式：{time_str}")
def loadMoreCleaver(AreaText):
    global StartText,breakText
    sText=AreaText[0]
    bText=AreaText[1]
    tempMsg=[] 
    needSetStartText= True if StartText==None else False
    needSetBreakText= True if breakText==None else False
    fastbreak=None#上一个上划中最小的时间，用于快速加载找到起始时间
    while(True) :
        print(f"开始找起始时间{fastbreak}")
        tempMsg=d(resourceId=versionWC.Time)
        print("开始找起始时间加载结束")
        canbreak=False 
        fastbreak=None
        for msg in tempMsg: 
            msgcontent=msg.get_text()
            timecontent=parse_wechat_time(msgcontent).date()
            if sText > timecontent or (needSetStartText==False and StartText==msgcontent):
                canbreak=True
            else:
                if(fastbreak==None):
                    fastbreak=msgcontent
                if(needSetStartText):
                    StartText=msgcontent
                    break
        if (needSetBreakText==True):
            for msg in (tempMsg): 
                msgcontent=msg.get_text()
                timecontent=parse_wechat_time(msgcontent).date()
                if bText!= timecontent:
                    if(needSetBreakText):
                        breakText=msgcontent
                    continue
                else:
                    needSetBreakText=False
                    print(f"*********找到结束时间了{breakText}")
                    break
        if(canbreak==True):
            print(f"*******找到起始时间了{StartText}") 
            break
        else:
            d(resourceId=versionWC.ControlList).swipe("down",10) 
            print("向上滚动")   
        # else:
        #     print("加载信息失败，可能没有更多信息了")
        #     return Exception("加载信息失败，可能没有更多信息了")
        #     if(fastbreak!=None and fastbreak==StartText):
        #         if(needSetStartText):
        #             StartText=None 
        #     break
    return tempMsg 
  
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
            if(len(dataread)<5 or timetohandle[0]=="" or (timetohandle[0]!="" and datetime.today().date()< datetime.strptime(timetohandle[0], "%Y/%m/%d").date())):
                endtimes.append((datetime.today()- timedelta(days=1)).date()) 
            else:
                endtimes=[datetime.strptime(datadate, "%Y/%m/%d").date() for datadate in timetohandle if datadate!=""]

    StartText=""#"昨天 21:15"# "2025年5月30日 3:14"#"0:25"#"2025年4月25日 5:48" 如果是None会根据配置自动找到开始统计的地方
    breakText=None#"0:12"#"星期二 17:00"#"昨天 9:10" #None#终止查询的时间节点6:44 如果是None会根据配置自动找到结束统计的地方
    DZDay=endtimes#点赞收藏的哪天 
    versionWC=WechatVersion("8.0.42") #微信版本号
    d = u2.connect() # 连接多台设备需要指定设备序列号 
    toinsertInfo1=[]
    cachesenders=[""]
    cachetimes=[""]
    cachecontents=[""]
    allin=False
    findtop=False
    tempValue=[]
    loadMoreCleaver((DZDay[0],DZDay[-1]))
    #----------------------找到聊天记录的最顶部---------------------------
    while(findtop==False):
        d(resourceId=versionWC.ControlList).swipe("down",10) 
        controlHoles=d(resourceId=versionWC.ControlHole)
        user33= controlHoles[0].child(resourceId=versionWC.User)[0].get_text()
        time33= DZDay[0]
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
            time33= DZDay[0]
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
                                messageset=content33[0].get_text()
                                if(sender!=""):
                                    mytempmessage=content33[0].get_text().split(sender)
                                    if(len(mytempmessage)>1):
                                        messageset=mytempmessage[1]
                                contenth=messageset+"@姜可艾 没有结算完"
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
 