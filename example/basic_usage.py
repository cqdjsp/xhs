import datetime
import json
from time import sleep

from playwright.sync_api import sync_playwright

from xhs.core import DataFetchError, XhsClient


def sign(uri, data=None, a1="", web_session=""):
    for _ in range(10):
        try:
            with sync_playwright() as playwright:
                stealth_js_path = "E:\\my\\job\\xhsTG\\public\\stealth.min.js"
                chromium = playwright.chromium

                # 如果一直失败可尝试设置成 False 让其打开浏览器，适当添加 sleep 可查看浏览器状态
                browser = chromium.launch(headless=False)

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


if __name__ == '__main__':
    cookie = "abRequestId=b1a8204b-f169-5ac9-a240-d7e76f92e284; a1=192bd6bf1cfstrjudljds60zw3ua7ycqcd1hniisp50000115806; webId=aa08832c525b96208379fb35dcbb81eb; gid=yj2yJ8S4q0SjyjJDfKDiyf33SiM9FV7f1VfMyMUK8uEq7x280WvSAI888yy2Y8K820iSyWdi; web_session=040069b73da8ec64049bdbd180354bb2cd3deb; webBuild=4.57.0; loadts=1740019516886; acw_tc=0a0d0f5817400196451531929e1550c0ac145544f1150a719390ab2ac1ecf2; websectiga=29098a4cf41f76ee3f8db19051aaa60c0fc7c5e305572fec762da32d457d76ae; sec_poison_id=b995ac0b-e4c1-4e0d-b56b-f8089453547f; x-user-id-creator.xiaohongshu.com=6649eba4000000000d0254cd; customerClientId=244158184254652; customer-sso-sid=68c517473327474294507721ead396574909caab; access-token-creator.xiaohongshu.com=customer.creator.AT-68c517473327474294336871falud8fmecdtl0kl; galaxy_creator_session_id=vkGmDW3N11pOMMAOEiFNwrTTAT3diXO8XfED; galaxy.creator.beaker.session.id=1740019646207039144021; xsecappid=creator-creator"

    xhs_client = XhsClient(cookie, sign=sign)
    print(datetime.datetime.now())

    for _ in range(10):
        # 即便上面做了重试，还是有可能会遇到签名失败的情况，重试即可
        try:
            note = xhs_client.get_note_by_id("67afecdf0000000028028c36","ABMuGHPzkrF3_R-x2Hv5gsctdOl93DbPpH4QcpptsADdg=")#,
            print(json.dumps(note, indent=4))
            print(help.get_imgs_url_from_note(note))
            break
        except DataFetchError as e:
            print(e)
            print("失败重试一下下")
