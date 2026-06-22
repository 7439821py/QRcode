import re
import time,random
import os
from datetime import datetime, timedelta
from playwright.sync_api import sync_playwright

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/147.0.7727.56 Safari/537.36"
)
登录信息 = r'C:\Users\Administrator\PycharmProjects\PythonProject\聚水潭成本计算\聚水潭胜算.json'
默认启动参数 = [
    "--disable-blink-features=AutomationControlled",  # 【最重要】移除 navigator.webdriver 特征
    "--no-sandbox",                                   # 防止权限报错（Linux/Docker 必备）
    "--disable-dev-shm-usage",                        # 防止内存崩溃
    "--disable-extensions",                           # 禁用扩展，减少变量暴露
    "--disable-background-networking",                # 减少后台噪音
    "--disable-default-apps",
    "--disable-sync",                                 # 禁用谷歌账号同步，减少特征
    "--no-first-run",
    "--disable-session-crashed-bubble",
    "--start-maximized",
]

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False,args=默认启动参数)
    context = browser.new_context(no_viewport=True,user_agent=UA,storage_state=登录信息)
    page = context.new_page()

    page.goto('https://ss.erp321.com/profit-report/multi-dimension')
    page.fill("input[placeholder='邮箱地址/手机号码']", "xxxxxx")
    page.fill("input[placeholder='请输入密码']", "xxxxxx")
    page.click('input[class="ant-checkbox-input"]')
    time.sleep(random.uniform(0.4,1))
    page.get_by_text('立即登录').click()
    time.sleep(random.uniform(0.4, 1))
    # goto 自带页面加载等待，前面的 sleep 多余
    page.goto('https://ss.erp321.com/profit-report/multi-dimension')

    page.get_by_text('高级筛选').click()

    page.locator('.ant-row:has(label:has-text("成本方案")) .ant-col.ant-form-item-control').click()
    # 等选项真的渲染出来再获取，比 sleep(1) 更可靠
    page.locator('div.ant-select-item-option-content').first.wait_for(state="visible", timeout=5000)
    成本列表对象 = page.locator('div.ant-select-item-option-content').all()
    成本列表 = [i.text_content() for i in 成本列表对象]
    # page.pause()
    # 成本列表=['分销成本']
    page.locator(".ant-drawer-title > div > .anticon > svg").click()  # 成本方案关闭


    def 日期选择(开始时间, 结束时间):
        target_s_month = 开始时间[:7]  # "2026-04"
        target_e_month = 结束时间[:7]  # "2026-06"

        page.get_by_role("textbox", name="自定义").click()

        # 定位上/下月按钮（请根据实际页面调整定位方式）
        上月按钮 = page.get_by_role("button").nth(3)
        下月按钮 = page.get_by_role("button").filter(has_text=re.compile(r"^$")).nth(4)

        def 获取当前年月():
            文本 = page.locator('div.ant-picker-header-view').nth(0).text_content()
            return datetime.strptime(文本, "%Y年%m月").strftime("%Y-%m")

        def 跳转到月份(目标年月):
            当前年月 = 获取当前年月()                 
            while 当前年月 != 目标年月:
                if 目标年月 < 当前年月:
                    上月按钮.click()
                else:
                    下月按钮.click()
                page.wait_for_timeout(300)
                当前年月 = 获取当前年月()

        # 选择开始日期
        跳转到月份(target_s_month)
        page.locator(f"td[title='{开始时间}']").click()

        # 选择结束日期
        跳转到月份(target_e_month)
        page.locator(f"td[title='{结束时间}']").click()

    # page.pause()
    page.locator('div.ant-form-item:has(label[title="退款统计口径"]) .ant-col.ant-form-item-control .ant-select').click() #找到锚点并选择右边的下拉框按钮
    page.locator('div.title:has-text("以确认时间统计")').hover()
    # page.locator('div.title:has-text("以确认时间统计")').highlight()
    time.sleep(1)
    page.locator('div.title:has-text("以确认时间统计")').click()

    今天 = datetime.today()
    开始时间 =今天.replace(day=1).strftime("%Y-%m-%d")
    昨天 = 今天 - timedelta(days=1)
    结束时间 = 昨天.strftime("%Y-%m-%d")

    日期选择(开始时间, 结束时间)

    now_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    today = datetime.now().strftime('%Y-%m-%d')

    # ── 统一导出函数：选择成本方案 → 选店铺 → 导出 ──────────
    def 选择店铺并导出(方案名, 店铺勾选列表):
        """
        店铺勾选列表:
          - "全选"         → 勾选「全选」后直接确认（公司成本）
          - [店铺名, ...] → 先取消全选，再逐个勾选列表中的店铺
        """
        page.get_by_text('高级筛选').click()
        page.locator('.ant-row:has(label:has-text("成本方案")) .ant-col.ant-form-item-control').click()
        page.locator(f'div.ant-select-item-option-content:has-text("{方案名}")').click()
        page.locator(".ant-drawer-title > div > .anticon > svg").click()  # 关闭抽屉

        page.get_by_role("textbox", name="店铺").click()
        page.get_by_role("tab", name="店铺部门归属").click()
        page.get_by_role("checkbox", name="展开所有").click()

        if 店铺勾选列表 == "全选":
            page.get_by_role("checkbox", name="全选").click()
            page.get_by_role("checkbox", name="利润(成本)中心[部门]").click()
        else:
            page.get_by_role("checkbox", name="全选").click()
            page.get_by_role("checkbox", name="全选").click()  # 点两次取消全选
            for 店铺名 in 店铺勾选列表:
                page.get_by_role("checkbox", name=店铺名, exact=True).click()
        time.sleep(1)
        page.get_by_role("button", name="确 定").click()

        page.get_by_role("button", name="查 询").click()
        page.locator('span.ss-spin-content-dot').wait_for(state="hidden", timeout=90000)
        page.get_by_role("button", name="导出数据").click()
        page.get_by_role("menuitem", name="导出明细数据").click()

        # 注册下载事件监听（不设超时，loading 结束后直接检查）
        捕获的下载 = []
        def on_download(download):
            捕获的下载.append(download)
        page.on("download", on_download)

        page.get_by_role("button", name="导 出").click()
        page.locator('button.ant-btn-loading').wait_for(state="hidden", timeout=320000)
        time.sleep(3)  # 给下载一点缓冲时间（此时监听器还在）
        page.remove_listener("download", on_download)

        if 捕获的下载:
            download = 捕获的下载[0]
            original_name = download.suggested_filename
            new_name = f"{方案名}{os.path.splitext(original_name)[1]}"
            folder_path = fr'C:\Users\Administrator\Desktop\聚水潭\{today}'
            os.makedirs(folder_path, exist_ok=True)
            save_path = os.path.join(folder_path, new_name)
            download.save_as(save_path)
            print(f"即时下载: {save_path}")
        else:
            # 没弹下载 → 尝试填备注 + 保存
            try:
                page.get_by_role("textbox", name="请输入备注").fill(f'{方案名} {now_time}', timeout=10000)
                time.sleep(1)
                page.get_by_role("button", name="保 存").click()
                备注列表.append(f'{方案名} {now_time}')
            except Exception as e:
                print(f"填备注失败: {e}")
                # 文本框不可用 → 暂停让你看页面状态
                # page.pause()


    page.pause()
    # ── 成本方案 → 店铺勾选 映射表 ──────────
    成本方案店铺映射 = {
        '公司成本':   "全选",
        '新媒体成本': ["新媒体", "新媒体-抖音", "新媒体-快手", "新媒体-视频号", "新媒体-小红书"],
        '三网成本':   ["三网-京东", "三网-拼多多", "三网-天猫",
                       "分销-陈瑶", "分销-林红发", "分销-马卫宽", "分销-鹏总",
                       "分销商-崔陈", "分销商-马卫宽", "分销-赵雨"],
        '分销成本':   ["分销-鹏总"],
    }

    备注列表 = ['三网成本5月','新媒体成本5月','公司成本5月']
    for 成本方案 in 成本列表:
        if 店铺列表 := 成本方案店铺映射.get(成本方案):
            选择店铺并导出(成本方案, 店铺列表)

    # ── 下载循环：逐个检查，下载成功就从待下载列表移除 ──
    待下载列表 = 备注列表.copy()
    while 待下载列表:
        page.get_by_role("button").filter(has_text=re.compile(r"^$")).nth(3).click()  # 打开任务列表
        print('任务列表已打开')
        for 备注 in 待下载列表.copy():  # 遍历副本，安全地 remove
            任务执行状态元素 = page.locator(f"//*[contains(., '{备注}')]/ancestor::tr//td[6]")
            任务执行状态 = 任务执行状态元素.text_content()
            print(任务执行状态)
            if '成功' in 任务执行状态:
                with page.expect_download() as download_info:
                    page.locator(
                        f"//*[contains(text(), '{备注}')]/ancestor::tr//div[@class='ant-space-item'][contains(., '下载')]"
                    ).click()
                download = download_info.value
            else:
                print(f"任务尚未完成: {任务执行状态}")
                original_name = download.suggested_filename
                方案 = 备注.split()[0]
                new_name = f"{方案}{os.path.splitext(original_name)[1]}"
                folder_path = fr'C:\Users\Administrator\Desktop\聚水潭\{today}'
                os.makedirs(folder_path, exist_ok=True)
                save_path = os.path.join(folder_path, new_name)
                download.save_as(save_path)
                print(f"文件已保存至: {save_path}")

                待下载列表.remove(备注)
                print(f"剩余待下载：{len(待下载列表)} 个")
                time.sleep(1)

        page.get_by_role("button", name="Close").click()
        if 待下载列表:
            print(f"等待 60 秒后重试剩余 {len(待下载列表)} 个任务…")
            time.sleep(60)

    # page.pause()
    browser.close()
