import uiautomator2 as u2
import logging
import sqlite3
import re
import time
import os
from datetime import datetime, timedelta

# ===================== 全局配置 & 常量 =====================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# 微信版本控件配置 新增版本只在这里加一行
WX_VERSION_CONF = {
    "8.0.42": {
        "User": "com.tencent.mm:id/lpa",
        "Content": "com.tencent.mm:id/lp8",
        "Time": "com.tencent.mm:id/lp_",
        "ControlList": "com.tencent.mm:id/lpg",
        "ControlHole": "com.tencent.mm:id/lp6",
        "ZFContent": "com.tencent.mm:id/cu2"
    },
    "8.0.60": {
        "User": "com.tencent.mm:id/lpa",
        "Content": "com.tencent.mm:id/lp8",
        "Time": "com.tencent.mm:id/lp_",
        "ControlList": "com.tencent.mm:id/lpg",
        "ControlHole": "com.tencent.mm:id/lp6",
        "ZFContent": "com.tencent.mm:id/cu2"
    },
    "8.0.70": {
        "User": "com.tencent.mm:id/gze",
        "Content": "com.tencent.mm:id/uls",
        "Time": "com.tencent.mm:id/tt7",
        "ControlList": "com.tencent.mm:id/lqa",
        "ControlHole": "com.tencent.mm:id/lp6",
        "ZFContent": "com.tencent.mm:id/cu2"
    }
}

# 业务常量
DB_REL_PATH = os.path.join("config", "WorkData.db")
SWIPE_DOWN_STEP = 10
SWIPE_UP_STEP = 75
WAIT_SECONDS = 3
NO_NEW_COUNT_LIMIT = 5
TOP_CHECK_COUNT = 4
BOUNDS_OFFSET = 28
TARGET_CONTENT_FLAG = "@姜可艾 没有结算完"

# 正则常量
PATTERN_TIME_FILTER = r'^\s*(?:(\d+)分(\d+)秒|(\d+\.?\d*)")\s*$'
PATTERN_UNDERLINE_TIME = r'(\d{4})_(\d{2})_(\d{2})_(\d{2})_(\d{2})_(\d{2})'

# ===================== 时间工具函数 =====================
def process_time(input_time: str) -> str | None:
    """
    多格式时间标准化为 YYYY-MM-DD HH:MM:SS
    异常返回None不崩溃
    """
    try:
        input_clean = input_time.strip()
        current_year = datetime.now().year
        yesterday = (datetime.now() - timedelta(days=1)).date()

        # 1. HH:MM
        m = re.match(r'^(\d{1,2}):(\d{2})$', input_clean)
        if m:
            h, mi = m.groups()
            h = int(h)
            mi = int(mi)
            return f"{yesterday} {h:02d}:{mi:02d}:00"

        # 2. HH:MM:SS
        m = re.match(r'^(\d{1,2}):(\d{2}):(\d{2})$', input_clean)
        if m:
            h, mi, s = map(int, m.groups())
            return f"{yesterday} {h:02d}:{mi:02d}:{s:02d}"

        # 3. MM月DD日 HH:MM
        m = re.match(r'^(\d{1,2})月(\d{1,2})日\s+(\d{1,2}):(\d{2})$', input_clean)
        if m:
            month, day, h, mi = map(int, m.groups())
            return f"{current_year}-{month:02d}-{day:02d} {h:02d}:{mi:02d}:00"

        # 4. MM月DD日 HH:MM:SS
        m = re.match(r'^(\d{1,2})月(\d{1,2})日\s+(\d{1,2}):(\d{2}):(\d{2})$', input_clean)
        if m:
            month, day, h, mi, s = map(int, m.groups())
            return f"{current_year}-{month:02d}-{day:02d} {h:02d}:{mi:02d}:{s:02d}"

        # 5. YYYY年MM月DD日 HH:MM:SS
        m = re.match(r'^(\d{4})年(\d{1,2})月(\d{1,2})日\s+(\d{1,2}):(\d{2}):(\d{2})$', input_clean)
        if m:
            year, month, day, h, mi, s = map(int, m.groups())
            return f"{year}-{month:02d}-{day:02d} {h:02d}:{mi:02d}:{s:02d}"

        raise ValueError("格式不匹配")
    except Exception as e:
        logger.error(f"时间解析失败: {input_time} | {str(e)}")
        return None
def extract_and_convert_time(input_str: str) -> datetime | None:
    """解析下划线分隔时间 2025_10_16_09_17_09"""
    try:
        m = re.search(PATTERN_UNDERLINE_TIME, input_str)
        if not m:
            raise ValueError("无匹配时间")
        year, month, day, hour, minute, second = map(int, m.groups())
        return datetime(year, month, day, hour, minute, second)
    except Exception as e:
        logger.error(f"下划线时间解析失败: {input_str} | {e}")
        return None

# ===================== 微信控件解析类 =====================
class WechatUiParser:
    def __init__(self, version: str):
        self.conf = WX_VERSION_CONF.get(version, WX_VERSION_CONF["8.0.42"])
        self.User = self.conf["User"]
        self.Content = self.conf["Content"]
        self.Time = self.conf["Time"]
        self.ControlList = self.conf["ControlList"]
        self.ControlHole = self.conf["ControlHole"]
        self.ZFContent = self.conf["ZFContent"]

# ===================== 核心爬取逻辑 =====================
def scroll_to_top(d: u2.Device, parser: WechatUiParser) -> bool:
    """滑动到聊天记录顶部"""
    temp_value = []
    find_top = False
    while not find_top:
        d(resourceId=parser.ControlList).swipe("down", SWIPE_DOWN_STEP)
        time.sleep(WAIT_SECONDS // 2)
        try:
            control_hole = d(resourceId=parser.ControlList)
            control_holes = control_hole.child()
            if control_holes.count < 1:
                continue

            first_item = control_holes[0]
            user = first_item.child(resourceId=parser.User)[0].get_text()
            t_time = first_item.child(resourceId=parser.Time)[0].get_text()
            content = first_item.child(resourceId=parser.Content)[0].get_text()

            current_tuple = (user, t_time, content)
            if not temp_value:
                temp_value.append(current_tuple)
            else:
                if temp_value[0] == current_tuple:
                    temp_value.append(current_tuple)
                else:
                    temp_value.clear()
                    temp_value.append(current_tuple)

            if len(temp_value) >= TOP_CHECK_COUNT:
                find_top = True
        except Exception as e:
            logger.warning(f"滑到顶部异常: {e}")
            continue
    logger.info("已到达聊天记录顶部")
    return True

def parse_one_page(d: u2.Device, parser: WechatUiParser, exist_set: set) -> list:
    """解析当前一页聊天记录，返回待入库数据"""
    page_data = []
    try:
        control_hole = d(resourceId=parser.ControlList)
        control_holes = control_hole.child(className="android.widget.LinearLayout")
        count = control_holes.count

        for i in range(count - 1):
            item = control_holes[i]
            content33 = item.child(resourceId=parser.Content)
            if not content33.exists:
                continue

            parent_bounds = item.bounds()
            sender = ""
            timeh = ""
            contenth = ""
            if content33.exists:
                bou = content33[0].bounds()
                if (bou[1] >= parent_bounds[1] and (bou[3] <= parent_bounds[3] or bou[3] - parent_bounds[3] < BOUNDS_OFFSET))==False:
                    continue
            else:
                continue
            # 解析发送人
            user33 = item.child(resourceId=parser.User)
            if user33.exists:
                bou = user33[0].bounds()
                if bou[1] >= parent_bounds[1] and bou[3] <= parent_bounds[3]:
                    sender = user33[0].get_text().strip()

            # 解析时间
            time33 = item.child(resourceId=parser.Time)
            if time33.exists:
                bou = time33[0].bounds()
                if bou[1] >= parent_bounds[1] and bou[3] <= parent_bounds[3]:
                    t_raw = time33[0].get_text().strip()
                    timeh = process_time(t_raw)

            # 解析内容
            if content33.exists:
                    
                raw_content = content33[0].get_text().strip()

                if not raw_content:
                    # 处理图片
                    if content33[0].info.get('contentDescription') == "图片":
                        continue
                    # 取转发内容
                    zf_item = item.child(resourceId=parser.ZFContent)
                    if zf_item.exists:
                        msg_set = zf_item[0].get_text().strip()
                        if re.match(PATTERN_TIME_FILTER, msg_set):
                            continue
                        if sender and sender in msg_set:
                            msg_set = msg_set.split(sender)[-1]
                        contenth = msg_set + TARGET_CONTENT_FLAG
                        contenth = contenth.replace(":", "")
                else:
                    if TARGET_CONTENT_FLAG not in raw_content:
                        continue
                    contenth = raw_content

            # 去重：三元组 发送人+时间+内容
            if contenth and timeh:
                key = (sender, timeh, contenth)
                if key not in exist_set:
                    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    page_data.append((sender, "[聊天]", "text", contenth, timeh, now_str))
                    exist_set.add(key)
    except Exception as e:
        logger.error(f"解析单页异常: {e}")
    return page_data

def crawl_chat(d: u2.Device, parser: WechatUiParser) -> list:
    """循环滑动爬取全部记录"""
    all_data = []
    exist_key_set = set()
    no_new_count = NO_NEW_COUNT_LIMIT

    while no_new_count > 0:
        time.sleep(WAIT_SECONDS)
        page_list = parse_one_page(d, parser, exist_key_set)
        if page_list:
            all_data.extend(page_list)
            no_new_count = NO_NEW_COUNT_LIMIT
            logger.info(f"当前页新增 {len(page_list)} 条记录")
        else:
            no_new_count -= 1
            logger.info(f"无新记录，剩余判定次数: {no_new_count}")

        # 上滑
        d(resourceId=parser.ControlList).swipe("up", SWIPE_UP_STEP)
        time.sleep(0.5)
    return all_data

# ===================== 数据库操作 =====================
def batch_insert_db(data_list: list):
    """批量插入数据库 自动事务+关闭"""
    if not data_list:
        logger.info("无数据需要入库")
        return

    # 确保文件夹存在
    os.makedirs(os.path.dirname(DB_REL_PATH), exist_ok=True)

    conn = None
    try:
        conn = sqlite3.connect(DB_REL_PATH)
        cur = conn.cursor()
        sql = '''
        INSERT INTO WXMSG (sender,myType,type,content,time,HandleDate)
        VALUES (?, ?, ?, ?, ?, ?)
        '''
        cur.executemany(sql, data_list)
        conn.commit()
        logger.info(f"成功入库 {len(data_list)} 条聊天记录")
    except Exception as e:
        if conn:
            conn.rollback()
        logger.error(f"数据库入库失败: {e}")
    finally:
        if conn:
            conn.close()

# ===================== 程序入口 =====================
def main():
    # 配置
    wx_version = "8.0.70"
    try:
        # 连接设备
        d = u2.connect()
        logger.info("设备连接成功")

        # 初始化解析器
        parser = WechatUiParser(wx_version)

        # 滑到顶部
        scroll_to_top(d, parser)

        # 爬取全部
        result_data = crawl_chat(d, parser)

        # 入库
        batch_insert_db(result_data)

        logger.info("爬取任务完成")
    except Exception as e:
        logger.error(f"程序运行异常: {e}", exc_info=True)

if __name__ == "__main__":
    main()