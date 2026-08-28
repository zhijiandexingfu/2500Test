# -*- coding: utf-8 -*-
"""
XGPS —— 中国移动香港「家居宽带 2500M(XGS-PON) 可售情况」批量排查工具
====================================================================

全局配置文件：
    * 接口地址 / 请求头 / 三个独立的「观镜防御系统」令牌 (XGiOG2f705)
    * 楼层采样规则（每 5 层随机选 1 层 curFloor + 当层随机选 1 个 flat curFlat）
    * 输入/输出文件路径
    * 结果表 (Excel) 列结构
    * 浏览器 DOM 选择器（自动化点击下拉第一项、选楼层单位、点击下一步）
    * 节流参数

修改配置优先级：命令行参数 > 本文件常量。
"""

import os

# ==================== 路径配置 ====================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 楼宇名称输入文件：每行一个楼宇名称 (inputBuilding)，支持以 # 开头的注释行
INPUT_FILE = os.path.join(BASE_DIR, "input_buildings.txt")

# 结果输出目录、Excel 落库文件、断点进度文件
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
EXCEL_PATH = os.path.join(OUTPUT_DIR, "result.xlsx")
PROGRESS_PATH = os.path.join(OUTPUT_DIR, "progress.json")   # 已完成楼宇集合（断点续跑）

# Playwright 浏览器持久化用户目录（保留 cookie，二次访问可跳过防御系统挑战）
BROWSER_PROFILE_DIR = os.path.join(BASE_DIR, ".browser_profile_xgps")

# 浏览器自动化执行 getInstallInfo 时所在的官网页面
BROADBAND_PAGE_URL = "https://www.hk.chinamobile.com/tc/home-family/broadband"

# ==================== 接口配置 ====================
BASE_URL = "https://www.hk.chinamobile.com"
API_SEARCH_ADDRESS = "/api/ecosp-emall/itemRest/homeBroadBand/searchAddress"
API_GET_ADDRESS_DETAIL = "/api/ecosp-emall/itemRest/homeBroadBand/getAddressDetail"
API_GET_INSTALL_INFO = "/api/ecosp-emall/itemRest/homeBroadBand/getInstallInfo"

# ---------------------------------------------------------------------------
# 「观镜防御系统」动态令牌：作为 URL 查询参数传递，参数名固定为 XGiOG2f705。
# 注意：三个接口的令牌【各自独立】（抓包时每个请求的令牌都不同）。
# 令牌会过期，过期后接口返回非 JSON / resCode != 000000；
# 可通过命令行 --token-search / --token-detail / --token-install 运行时覆盖。
# ---------------------------------------------------------------------------
TOKEN_PARAM = "XGiOG2f705"

# searchAddress 专用令牌（前台浏览器自动化点击搜索时随 URL 自动带上；
#                       同时也是兜底——若浏览器捕获失败，可由后台 requests 复用）
XGIOG_TOKEN_SEARCH = (
    "bfo569818597EsndR6zX76117040b81.TJCZ2RF1Nm8fni3sVAXS0CYKt7XkAyNogJ9DrzDsfBUm2na"
    "WuCSsB6b5006bg.CVoJBwe4qyHboTflEtLAKxKYd6fzftrBOlwipVhw39bq22mIHkl5Hr._2CjJ.PZx"
    "dQKucf2OWj1qADVLHFMyfGjRJlA2wZmDq6kwTHjCPFCrMtiz.go.ee7W7TbPwMDGJSmG13IhTgDmXIh"
    "wP3hw0_ZYXvI.DjrgZqApqkYwFAWY4ebo5vofYfoYW.fbPM0cvX4JsCpejpl5f2IFAwBWCyQvONiuAf"
    "jZ6_iqxBdL4"
)

# getAddressDetail 专用令牌（后台 requests 调用）
XGIOG_TOKEN_DETAIL = (
    "beo1775563852xb4Bcxfk21d053e6Utg3GE3PpZw9x7jNqwnWFtHpD3WXB2yzL73tqwOclc_HiP0P.D"
    ".30LMQvXqw5Z.p.4LqwRTUi9.F12LpakMPFvIBPx8MJ90BVIG0vXY6LP9J.ntnBpk3UuJ9giO2_lsMX"
    "w_1ocbL0v_slalX.rUrXyfj0d6onzUF0WHKq8nhiBu6c3ClxU1kHD1QKz6T08R3JVzNlNa.qsCwpRn7k"
    "a3hfj.CPrlIZ69Gz1ocd2li5CoPL0Pu9MtDH2Dzl0FimGZ2bQV6rb.JGHlXGuF90txmTY7F4Ga6LrzL"
    "9YJSopthfw4"
)

# getInstallInfo 专用令牌（浏览器页面内 fetch 调用，模拟点击下一步后的接口请求）
XGIOG_TOKEN_INSTALL = (
    "baj919709285wT63Dfp87f9e1532V7J_mLw7Tb7xhYcvaw942Hitp63K3Od1WkA3ulPnxrHlGKbL7j2Q"
    "fZhU0Tdh_iNgl2XtOE0nqupSWArCemqntQvEz3I_.YJs3nbicz0E2a9OVdneTkBLBUOVnbcxpoCEY6uG"
    "gPmmBiqF2Tt4rR6D68p0EK3_P.luUZkRaQli5gFqO03NPO6XKhLZgfIyvrr2P3fUeZfOdJBCCjfocfwE"
    "dj2zepLCDoz1sRL9wgZFnlZjLpwhWhFnVBDLC4ZhmZeudQDShG8aeBh8JjZAxpx_uba0aZowITni8_BO"
    "RU2Ic7Q"
)

# 前端埋点设备号（抓包自官网请求头，保持与浏览器一致可降低风控风险）
OMNI_CLIENT_DEVICE_ID = "OCD0202603132032360919656300544"

# 后台 requests 公共请求头（从官网抓包还原）
DEFAULT_HEADERS = {
    "acc-lang": "zh-HK",
    "accept": "application/json, text/plain, */*",
    "accept-language": "zh-HK",
    "api-version": "2",
    "channelid": "WWW",
    "cmhkchannel": "WWW",
    "content-type": "application/x-www-form-urlencoded; charset=UTF-8",
    "omni-client-device-id": OMNI_CLIENT_DEVICE_ID,
    "origin": BASE_URL,
    "referer": BROADBAND_PAGE_URL,
    "user-agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/150.0.0.0 Safari/537.36"
    ),
}

# 浏览器页面内 fetch 还需要显式设置的业务请求头（同源，浏览器自动带 cookie + UA）
# 这里只列出关键几个，其他由浏览器自动带
BROWSER_FETCH_HEADERS = {
    "accept": "application/json, text/plain, */*",
    "api-version": "2",
    "channelid": "WWW",
    "cmhkchannel": "WWW",
    "content-type": "application/x-www-form-urlencoded; charset=UTF-8",
    "omni-client-device-id": OMNI_CLIENT_DEVICE_ID,
}

# ==================== 浏览器 DOM 选择器 ====================
# 探测自官网宽带页（见 output/probe4_find_real_selector.py），
# Element UI 在这个版本里用了自定义类名，不要硬编码历史版本。
SELECTORS = {
    # 搜索框
    "search_input": "#searchInput",
    # 联想下拉的地址项（每项是一个 .searchGoodsLine）
    "dropdown_items": ".broadband-search .searchGoodsLine",
    # 选楼层/单位：第一个 .el-select 是楼层，第二个是单位
    "floor_select": ".el-select >> nth=0",
    "flat_select": ".el-select >> nth=1",
    "select_dropdown_item": ".el-select-dropdown:visible .el-select-dropdown__item",
    # 下一步按钮（前置步骤页面，ID 来自真实 DOM）
    # 主选择器用 #btn_family_001；fallback 兼容老版本
    "next_button": "#btn_family_001",
    "next_button_fallback": "button:has-text('下一步')",
    # 入口 / 进度 段，常用于判断"已回到搜索步骤"
    "search_section": ".broadband-search",
    # 结果页「重新查詢」按钮：点击后回退到地址搜索步骤（省整页重载）
    "requery_button": "#btn_family_003",
}

# ==================== 业务参数 ====================
CLIENT_TYPE = "STANDALONE"   # 客户类型（住宅独立客户）
BROADBAND_TYPE = "CMHK"      # 默认查询 CMHK 自营宽带
REQUEST_TIMEOUT = 20         # 单次 HTTP 请求超时（秒）

# ==================== 楼层采样规则 ====================
# 规则：把楼宇所有楼层排序后按每 FLOOR_GROUP_SIZE 层分组，
#       每组随机选 1 层记为 curFloor，再在该层的所有 Flat 中随机选 1 个记为 curFlat。
#       即一栋 30 层的楼会产出 6 组采样（每组一条 Excel 记录）。
FLOOR_GROUP_SIZE = 5

# 采样随机种子前缀：种子 = SEED_PREFIX + inputBuilding。
# 使用固定种子使「同一楼宇多次运行的采样结果可复现」，保证断点续跑时行键对齐。
SEED_PREFIX = "XGPS|"

# ==================== 结果表结构 ====================
# 基础列（题目要求的固定表头）
BASE_HEADERS = [
    "inputBuilding",   # 输入的楼宇名称
    "curBuilding",     # searchAddress 返回的第一个楼宇地址
    "curFloor",        # 采样楼层（每 5 层随机选 1 层）
    "curFlat",         # 采样单位（该楼层随机选 1 个 Flat）
    "Is2500Support",   # 是否可卖 2500M：Y / N / NA
    "defeatBuilding",  # 是否为缺陷楼宇：Y / N / NA
    "remark",          # 备注：可卖2500M / 不可卖2500M / ...
]

# 扩展列（辅助 bad case 分析，可按需增删，程序自动适配）
EXTRA_HEADERS = [
    "buildingCode",      # 楼宇编码（CMHK buildingCode）
    "ofcaCode",          # OFCA 编码
    "carrier",           # 覆盖运营商（getInstallInfo 返回）
    "isXGSPONsupport",   # 接口侧 XGS-PON(2500M) 支持标识（浏览器内调用 getInstallInfo 获得）
    "coverType",         # 覆盖类型（如 GP）
    "floorGroup",        # 采样分组（如 3-7，表示该行采自第 3~7 层这一组）
    "checkedAt",         # 本行采集时间
    "level11_address",   # 11级地址，等价于 curBuilding,curFloor + "层" + curFlat + "户"
]

# ==================== 接口结果 -> 判定列 自动映射 ====================
# isXGSPONsupport 为接口侧 2500M(XGS-PON) 支持标识，直接映射到判定三列
AUTO_RESULT_MAP = {
    "Y": {"Is2500Support": "Y", "defeatBuilding": "N", "remark": "可卖2500M"},
    "N": {"Is2500Support": "N", "defeatBuilding": "Y", "remark": "不可卖2500M"},
}

# ==================== 节流配置 ====================
API_CALL_INTERVAL = 0.8   # 同一楼宇内两次接口调用的间隔（秒），避免触发风控
REQUEST_INTERVAL = 1.5    # 相邻楼宇之间的间隔（秒）
