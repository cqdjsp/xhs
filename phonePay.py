import xlwings as xw
import uiautomator2 as u2
import time
import logging
import random
import sys
import os

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class WechatVersion:
    def __init__(self, version):
        """初始化微信版本检查"""
        self.version = version
        if(version=="8.0.42"):
            self.Plus="com.tencent.mm:id/ky9"#加号按钮 
            self.Picture="com.tencent.mm:id/hnn"#相册按钮
            self.SelectPicture="com.tencent.mm:id/je0"#选择图片按钮
            self.AmountInput="com.tencent.mm:id/pbn"#金额输入框
            self.OtherAmount="com.tencent.mm:id/lgk"#其他金额按钮
            self. ConfirmAmount="com.tencent.mm:id/lfv"#赞赏码确认金额按钮
            self.PayButton="com.tencent.mm:id/hql"#支付码确认金额按钮付款按钮
            self.ZSperson="com.tencent.mm:id/lfq"#赞赏人姓名
            self.PayPassword="com.tencent.mm:id/tenpay_keyboard_"# # 支付密码键盘 
            self.PaySuccess="com.tencent.mm:id/jla"#不开通指纹支付按钮
        elif(version=="8.0.54"):
            self.Plus="com.tencent.mm:id/plus_icon"#加号按钮 
            self.Picture="com.tencent.mm:id/hnn"#相册按钮
            self.SelectPicture="com.tencent.mm:id/je0"#选择图片按钮
            self.AmountInput="com.tencent.mm:id/pbn"#金额输入框
            self.OtherAmount="com.tencent.mm:id/lgk"#其他金额按钮
            self. ConfirmAmount="com.tencent.mm:id/lfv"#赞赏码确认金额按钮
            self.PayButton="com.tencent.mm:id/keyboard_action"#支付码确认金额按钮付款按钮
            self.ZSperson="com.tencent.mm:id/lfq"#赞赏人姓名
            self.PayPassword="com.tencent.mm:id/tenpay_keyboard_"# # 支付密码键盘 
            self.PaySuccess="com.tencent.mm:id/jla"#不开通指纹支付按钮
class WeChatDonation:
    def __init__(self, excel_path, password=None,startindex=0,versionWC=WechatVersion("8.0.42")):
        """初始化赞赏助手，加载Excel数据"""
        self.excel_path = excel_path
        self.password = password  # 支付密码，如需要
        self.shtPayDetail = None  # 支付详情工作表
        self.wb=None
        self.data = self.load_excel()
        self.d = None  # uiautomator2设备对象
        self.startindex=startindex
        self.OBJWC=versionWC
    def load_excel(self):
        """从Excel读取赞赏码和金额数据"""
        try:
            wb = xw.Book(self.excel_path)
            sheet=wb.sheets["Sheet1"]  # 假设数据在第一个工作表
            # 确保列名包含"赞赏码"和"金额"
            row_values = sheet.range('B1').value
            row_values2 = sheet.range('D1').value
            required_columns = ["备注号", "按小红书查到的计算"] 
            if row_values not in required_columns or row_values2 not in required_columns:
                raise ValueError(f"Excel中缺少必要列")
            self.wb=wb
            for sheettemp in wb.sheets:
               if(sheettemp.name=="支付详情"):
                   self.shtPayDetail =sheettemp
                   break
            if   self.shtPayDetail==None:
                self.shtPayDetail =wb.sheets.add(name='支付详情')
                self.shtPayDetail.range(f'A{1}').value = list(("微信名", "备注号", "按小红书查到的金额","支付微信号","支付金额"))
            return sheet.range('A2').expand().value
        except Exception as e:
            logger.error(f"读取Excel失败: {str(e)}")
            sys.exit(1)
    
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
    
    def open_scanner(self):
        """打开微信扫一扫"""
        try:
            # 点击加号按钮
            if not self.d(resourceId=self.OBJWC.Plus).exists(timeout=2):
                return False
            self.d(resourceId=self.OBJWC.Plus).click()
            time.sleep(0.5)  # 等待菜单弹出
            if self.d(text="扫一扫").exists(timeout=2):
            # 点击扫一扫
                self.d(text="扫一扫").click()
            return True
        except Exception as e:
            logger.error(f"打开扫一扫失败: {str(e)}")
            return False
    
    def scan_qrcode(self, qrcode,amount,row):
        """扫描赞赏码"""
        try:
            qrcode_path=f"config/zfcode/{qrcode}.jpg"
            logger.info(qrcode_path)
            #qrcode_path=f"config/zfcode/1.jpg"
            upload_image(self.d,qrcode_path, "/sdcard/Pictures/WeiXin/1.jpg")
            # 点击相册按钮
            time.sleep(random.uniform(0.3, 0.6))
            if(self.d(resourceId=self.OBJWC.Picture).exists(timeout=2)):
                self.d(resourceId=self.OBJWC.Picture).click()
            time.sleep(random.uniform(0.3, 0.6))
            if(self.d(resourceId=self.OBJWC.SelectPicture).exists(timeout=2)):
                time.sleep(random.uniform(1.5, 2))
                self.d(resourceId=self.OBJWC.SelectPicture).click() 
            """输入支付金额"""
            isZan=True
            userName=None
            if self.d(resourceId=self.OBJWC.OtherAmount).exists(timeout=6):
               #使用的是赞赏码
               if not self.d(resourceId=self.OBJWC.OtherAmount).exists(timeout=10):
                    return (False,None)
               else:
                    time.sleep(random.uniform(0.3, 0.6))
                    self.d(resourceId=self.OBJWC.OtherAmount).click()#点击其他金额  
                    # 等待金额输入框出现
                    if not self.d(resourceId=self.OBJWC.AmountInput).exists(timeout=5):
                        logger.error("未找到金额输入框")
                        return (False,None)
                    logger.info("扫描赞赏码成功")
                    userName=self.d(resourceId=self.OBJWC.ZSperson).get_text()  # 获取赞赏人姓名
                    # 输入金额
                    self.d(resourceId=self.OBJWC.AmountInput).set_text(str(amount))
                    if not self.d(resourceId=self.OBJWC. ConfirmAmount).exists(timeout=2):
                        d.swipe_ext("up", scale=0.5)
                    time.sleep(random.uniform(0.3, 0.5))
                    logger.info("输入金额成功")    
                    if not self.d(resourceId=self.OBJWC. ConfirmAmount).exists(timeout=1):
                        logger.error("未找到确认金额按钮")
                        return (False,None)                       
                    self.d(resourceId=self.OBJWC. ConfirmAmount).click() 
                    logger.info("点击赞赏成功") 
            else:#支付码
                isZan=False
                try:
                    # 等待金额输入框出现
                    if not self.d(resourceId=self.OBJWC.AmountInput).exists(timeout=5):
                        logger.error("未找到金额输入框")
                        return (False,None)
                    logger.info("扫描支付码成功")
                    
                    # 输入金额
                    self.d(resourceId=self.OBJWC.AmountInput).set_text(str(amount))
                    time.sleep(random.uniform(0.3, 0.5))
                    logger.info("输入金额成功")
                    # 点击"确定"按钮
                    self.d(resourceId=self.OBJWC.PayButton).click()
                    time.sleep(random.uniform(0.3, 0.5))
                    logger.info("点击付款成功")
                     
                except Exception as e:
                    logger.error(f"输入金额失败: {str(e)}")
                    return (False,None)                
        except Exception as e:
            logger.error(f"扫描赞赏码失败: {str(e)}")
            return (False,None)
        """确认支付"""#识别并支付 
        try:
            # if   self.d(text="继续支付").exists(timeout=2):
            #     self.d(text="继续支付").click()
            # if   self.d(text="我知道了").exists(timeout=2):
            #     self.d(text="我知道了").click()

            # 等待支付确认界面 
            myhandlevalue=None
            if(isZan):
                if not self.d(text="请输入支付密码").exists(timeout=3):
                    logger.error("未找到请输入支付密码提示") 
                logger.info(f"正在处理赞赏给{userName}的赞赏{amount}元")
                myhandlevalue=list((row[0],int(row[1]),row[4],userName,str(amount)))
            else:
                if   self.d(text="识别并支付").exists(timeout=2):
                    self.d(text="识别并支付").click()
                if not self.d(text="请输入支付密码").exists(timeout=3):
                    logger.error("未找到请输入支付密码提示")
                      
                logger.info(f"正在处理支付给{row[0]}的支付{amount}元")
                textviews = d.xpath('//android.widget.TextView').all()
                for i, textview in enumerate(textviews) :
                    if("付款给" in textview.text): 
                        logger.warning(f'处理{textview.text}的支付{textviews[i+1].text}元') 
                        myhandlevalue=list((row[0],int(row[1]),row[4],textview.text,textviews[i+1].text))
                        #self.shtPayDetail.range(f'A{i+2}').value = textview.text
                        break
            if self.password:
                trycount = 0# 尝试3次输入密码
                # 如果设置了密码，自动输入 
                i=0
                while(i<len(self.password)) :  
                    
                    digit  = self.password[i]
                    self.d(resourceId=f"{self.OBJWC.PayPassword}{digit}").click()
                    i += 1
                    time.sleep(random.uniform(0.3, 0.5)) 
                    if(self.d(resourceId=self.OBJWC.AmountInput).exists(0.3) and len(self.d(resourceId=self.OBJWC.AmountInput).get_text())!=i):
                        logger.warning(f"{digit}输入失败")
                        i -= 1  
                        trycount+=1
                    else: 
                        trycount = 0
                      # 获取当前输入的密码
                    if trycount >= 3:   
                        logger.error("输入支付密码失败")
                        break
            else:
                # 等待用户手动输入密码
                logger.info("请手动输入支付密码...")
                time.sleep(10)  # 给用户10秒时间输入密码
            if self.d(resourceId=self.OBJWC.PaySuccess).exists(timeout=2):
                self.d(resourceId=self.OBJWC.PaySuccess).click()  # 稍后再说   不开指纹支付
            # 检查是否支付成功
            if self.d(text="支付成功").exists(timeout=5):
                logger.info("支付成功!")
                self.d(text="完成").click() 
            else:
                logger.warning("未检测到支付成功提示，可能需要手动确认") 
            qrcode_path="/sdcard/Pictures/WeiXin/1.jpg"
            self.d.shell(f"rm {qrcode_path}")
            return (True,myhandlevalue)
           
        except Exception as e:
            logger.error(f"确认支付失败: {str(e)}")
            return (False,None)
    
    def process_payments(self):
        """处理所有支付"""
        if not self.connect_device():
            return False
        
        if not self.open_wechat():
            return False
        
        success_count = 0
        total_count = len(self.data)
        startindex=self.startindex
        for index, row in enumerate(self.data):
            if index < startindex:
                continue
            qrcode = int(row[1])#"", ""
            amount = row[4]
            wxname= row[0]
            logger.info(f"开始处理 {qrcode}  {wxname}的支付，金额: {amount}")
            if(amount>0):
                if self.open_scanner():
                    returnvalue=self.scan_qrcode(qrcode,amount,row) 
                    if   returnvalue[0]: 
                        success_count += 1
                        self.shtPayDetail.range(f'A{index+2}').value = returnvalue[1]
                        # 支付完成后返回主界面
                        #self.back_to_main()
            else:
                self.shtPayDetail.range(f'A{index+2}').value=list((wxname,qrcode,amount))
            # 每笔支付后稍作休息，避免操作过快
            time.sleep(random.uniform(0.3, 1))
        self.wb.save()
        self.wb.close()
        logger.info(f"支付任务完成! 总共 {total_count} 笔，成功 {success_count} 笔")
        return success_count == total_count

def upload_image(d, local_path, device_path):
    """上传图片到设备""" #识别并支付

    try:
        d.push(local_path, device_path)
        d.shell(f"chmod 664 {device_path}")
        d.shell(f"am broadcast -a android.intent.action.MEDIA_SCANNER_SCAN_FILE -d file://{device_path}")
        logger.info(f"图片已上传并设置权限: {device_path}")
    except Exception as e:
        logger.error(f"上传图片失败: {str(e)}")
if __name__ == "__main__":
    # Excel文件路径，确保文件存在且格式正确
    excel_path = "E:/my/job/xhs/Result/78.xls"  
    startindex=1#excel表格的行号-2
    versionWC=WechatVersion("8.0.54") #微信版本号
    d = u2.connect() # 连接多台设备需要指定设备序列号
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
        donation.wb.close()