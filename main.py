# -*- coding: utf-8 -*-
"""
XGPS 主程序 —— 中国移动香港家居宽带 2500M(XGS-PON) 可售情况批量排查
====================================================================

执行流程（按用户在最新提示词中的「浏览器自动化 + 后台辅助」混合流程）：
    1.【浏览器】打开宽带页 -> 在 #searchInput 输入地址 -> 点击下拉第一项
                -> 同时浏览器自动触发 searchAddress，并由 response 监听器捕获
    2.【后台 requests】用捕获到的地址对象调 getAddressDetail
                -> 返回值 busiResp 保存为 busiRespObj，从中取出所有 floor/flat
    3.【采样】每 5 层随机选 1 层 (curFloor) + 当层随机选 1 个 Flat (curFlat)
                用楼宇名做随机种子，保证重跑结果一致（断点续跑行键对齐）
    4.【浏览器】等待页面进入"选楼层/单位"步骤
                可视化点击楼层 + 单位两个下拉（让用户在浏览器里看到操作）
    5.【人工判定（默认 --judge）】点「下一步」#btn_family_001 -> 跳到结果页
                -> 注入浮动判定面板（可卖2500M/不可卖2500M/其他情况待排/跳过）
                -> 等用户【终端输入 Y/N/E/S】或【浏览器点按钮】，无输入则一直等
       或【全自动 --auto】在页面同源上下文里 fetch getInstallInfo
                (走观镜防御系统放行后的浏览器通道，使用专用令牌)
                -> 取 isXGSPONsupport 字段映射到 Is2500Support / defeatBuilding
    6.【落库】每个采样一组一行实时写入 Excel（行键幂等），支持楼宇级+行级断点续跑

判定映射（人工模式）：Y=可卖2500M；N=不可卖2500M(缺陷楼宇)；E=异常待排；S=跳过

用法示例：
    python main.py                                     # 人工判定模式（默认）
    python main.py --auto                              # 全自动模式（fetch 自动判定）
    python main.py --limit 3                           # 只跑前 3 栋
    python main.py --redo                              # 忽略已有结果，全部重跑
    python main.py --headless                          # 无头浏览器（可能被风控拦截，慎用）
    python main.py --token-search <t> ...              # 令牌过期时按接口分别注入新值
"""

from __future__ import annotations

import argparse
import logging
import queue
import random
import re
import sys
import threading
import time
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

import config
from api_client import ApiError, CmhkApiClient
from browser_bot import BrowserBot
from storage import ExcelStorage, ProgressTracker


logger = logging.getLogger("xgps.main")


# ----------------------------------------------------------------------
# 地址匹配校验（防止模糊搜索"张冠李戴"）
# ----------------------------------------------------------------------
# CMHK searchAddress 是模糊搜索：输入地址不存在时会返回一堆近似候选，
# 旧逻辑无脑点下拉第一项，会把别的楼宇（如输入"新科技广场"却解析成
# "大有大廈"）当成结果。此校验在采样前确认系统返回的地址确实包含
# 输入楼宇名，否则直接写「地址未匹配」降级行。

try:                                   # opencc 为可选依赖：简体输入自动转繁体再比对
    from opencc import OpenCC
    _S2T = OpenCC("s2t")

    def _to_traditional(text: str) -> str:
        return _S2T.convert(text)
except Exception:                      # 缺库时回退原文比对（输入本身是繁体则无影响）
    def _to_traditional(text: str) -> str:
        return text


# 座号 token / 座号前缀：第2座 / A座 / 12座，以及「第4座海韻花園」粘连形式
_BLOCK_TOKEN_RE = re.compile(r"^(?:第?[0-9A-Za-z一二三四五六七八九十]+座)$")
_BLOCK_PREFIX_RE = re.compile(r"^第?[0-9A-Za-z一二三四五六七八九十]+座")


def extract_building_name(keyword: str) -> str:
    """
    从逗号分隔地址中取楼宇名。
    地址格式约定为「[座,]楼宇,街道,地区」—— 取第一个非座号 token，
    并剥离「第4座海韻花園」这类座号粘连前缀。
    """
    tokens = [t.strip() for t in keyword.split(",") if t.strip()]
    cleaned = [_BLOCK_PREFIX_RE.sub("", t) for t in tokens]
    cleaned = [t for t in cleaned if t and not _BLOCK_TOKEN_RE.match(t)]
    if not cleaned:
        cleaned = tokens or [keyword.strip()]
    return cleaned[0]


def is_address_matched(keyword: str, cur_building: str) -> bool:
    """校验系统解析到的地址包含输入楼宇名（简繁归一化后子串匹配）。"""
    name = _to_traditional(extract_building_name(keyword))
    target = _to_traditional(cur_building or "")
    return bool(name) and name in target


# ----------------------------------------------------------------------
# 人工判定：主通道 = 浏览器内键盘/点击 + fallback = 终端 stdin（EOFError→E）
# ----------------------------------------------------------------------
def wait_for_judgment(bot: BrowserBot, prompt: str) -> str:
    """
    等待用户判定。

    主通道（浏览器获焦）：
      - 在结果页按 Y / N / E / S → 浮动面板自动判定
      - 点击浮动面板 4 按钮（可卖/不可卖/待排/跳过）→ 自动判定

    Fallback 通道（浏览器意外死亡 / polling 抛异常）：
      - 终端提示明确，返回 fallback 提示
      - 捕获 EOFError 自动降级为 "E"

    没有任何输入会一直轮询；Ctrl+C 向上抛出 KeyboardInterrupt。
    """
    valid = ("Y", "N", "E", "S")
    # 一次性打印，避免循环里刷屏
    print(prompt, end="", flush=True)

    poll_errors = 0
    while True:
        # ---- 主通道：浏览器判定面板 / 键盘事件 ----
        try:
            val = bot.read_judge_click()
            poll_errors = 0  # 成功一次就清零累计
            if val in valid:
                print(f" [浏览器判定] {val}", flush=True)
                # 用完即清空，避免下一组残留（inject_judge_panel 也会清，这里兜底）
                try:
                    bot.reset_judge_click()
                except Exception:
                    pass
                return val
        except Exception as exc:
            poll_errors += 1
            logger.warning("浏览器判定通道异常 (%d/3): %s", poll_errors, exc)
            if poll_errors >= 3:
                # 浏览器会话死透了，回退到终端 input()
                print(
                    "\n    [浏览器会话异常] fallback 到终端输入；EOFError 将降级为 E",
                    flush=True,
                )
                try:
                    ans = input("    判定 (Y/N/E/S，EOF→E): ").strip().upper()
                except EOFError:
                    print("    [EOFError] 降级为 E")
                    return "E"
                if ans in valid:
                    return ans
                return "E"  # 任何非法输入都降级为 E（不卡）

        time.sleep(0.25)


# ----------------------------------------------------------------------
# 工具函数
# ----------------------------------------------------------------------
def read_input_lines(path: str) -> List[str]:
    """读取楼宇名称输入文件：每行一个名称，跳过空行与 # 注释行。"""
    with open(path, "r", encoding="utf-8-sig") as f:   # utf-8-sig 兼容带 BOM 的记事本文件
        lines = [ln.strip() for ln in f]
    return [ln for ln in lines if ln and not ln.startswith("#")]


def pick_building_code(carrier_info: List[Dict[str, Any]]) -> str:
    """从 carrierInfo 中提取 CMHK 楼宇编码（priority=1 那条）。"""
    if not carrier_info:
        return ""
    items = sorted(carrier_info, key=lambda c: c.get("priority", 99))
    return items[0].get("addressCode", "")


def build_install_bus_info(
    address_obj: Dict[str, Any],
    building_code: str,
    floor: str,
    flat: str,
) -> Dict[str, Any]:
    """
    构造 getInstallInfo 的 busInfo 入参。约定与官网前端一致。
    注意：carrierInfo 要和 searchAddress 返回里的原始结构一致
          （保留 addressCode 字段）。
    """
    carrier_info = address_obj.get("carrierInfo") or []
    return {
        "broadBandType": config.BROADBAND_TYPE,
        "buildingCode": building_code,
        "clientType": config.CLIENT_TYPE,
        "ofcaCode": address_obj.get("ofcaCode", ""),
        "carrierInfo": carrier_info,
        "floor": str(floor),
        "flat": str(flat),
    }


def _numeric_then_lex_key(value: Any):
    """排序键：先尝试转 int 比较（自然数序），失败则用原值的字符串字典序兜底。

    用于楼层/单位这种「既能是数字又能是非数字」的样本，确保：
      - 纯数字楼层（如 1, 2, 10）按 1 < 2 < 10 排序，而不是 "10" < "2"
      - 含中文/字母的兜底用 Unicode 字符串字典序
      - 大写/小写字母归一（视工程需要可选；这里默认保留原样）
    """
    s = str(value)
    try:
        return (0, int(s), s)   # 第一段相等表示都是数字，按 int 排，再附原串作稳定 tie-breaker
    except (TypeError, ValueError):
        return (1, s, "")       # 非数字分组排后面，按字符串字典序


def pick_sampled_units(
    busi_resp_obj: Dict[str, Any],
    seed: str,
    group_size: int = config.FLOOR_GROUP_SIZE,
) -> List[Tuple[str, str, str]]:
    """
    从 getAddressDetail 的 busiRespObj 提取该楼宇所有 floor/flat，
    先按数字（int）排序、相同 floor 按 Flat 字典序排序，
    再按 group_size 分组，每组随机选 1 层 curFloor，再在该层随机选 1 个 Flat curFlat。

    返回: [(curFloor, curFlat, "分组范围如3-7")]，列表为空表示无楼层数据。
    """
    data = busi_resp_obj.get("busiDataResp") or {}
    units = data.get("homeUnitList") or []
    if not units:
        return []

    # 每层：先把 flats 字典序排好（chunk 内随机抽也只挑「排序后的稳定列表」）
    #     再把每层打包成 (floor, flats_sorted) 用于楼层排序
    floors: List[Tuple[Any, List[Any]]] = []
    for u in units:
        f = u.get("floor")
        flats = u.get("flat") or []
        if f is None or not flats:
            continue
        sorted_flats = sorted(flats, key=_numeric_then_lex_key)
        floors.append((f, sorted_flats))

    # 楼层排序：先按数字序（int 升序，non-int 落到后面），相同 floor 走兜底字符串
    floors.sort(key=lambda x: _numeric_then_lex_key(x[0]))

    if not floors:
        return []

    # 固定种子绑定楼宇名，采样结果可复现（断点续跑行键对齐）
    rng = random.Random(f"{config.SEED_PREFIX}{seed}")

    samples: List[Tuple[str, str, str]] = []
    for i in range(0, len(floors), group_size):
        chunk = floors[i : i + group_size]
        if not chunk:
            continue
        chosen_floor, chosen_flats = rng.choice(chunk)
        chosen_flat = rng.choice(chosen_flats)
        # 组范围（用楼层序号表达，不一定连续）
        lo, hi = chunk[0][0], chunk[-1][0]
        group_label = f"{lo}-{hi}"
        samples.append((str(chosen_floor), str(chosen_flat), group_label))
    return samples


def summarize_install(install_data: Dict[str, Any]) -> Dict[str, str]:
    """
    把 getInstallInfo 返回提炼为 Excel 记录字段。

    :param install_data: 完整 busiResp JSON（包含 busiResp 包装层）
    """
    busi = (install_data or {}).get("busiResp") or {}
    install = busi.get("busiDataResp") or {}
    cover = install.get("coverType")
    if isinstance(cover, list):
        cover = ",".join(str(c) for c in cover)
    return {
        "isXGSPONsupport": str(install.get("isXGSPONsupport", "") or ""),
        "carrier": str(install.get("carrier", "") or ""),
        "coverType": str(cover or ""),
    }


def apply_auto_result(record: Dict[str, Any]) -> Dict[str, Any]:
    """按接口 isXGSPONsupport 自动填充判定三列。"""
    mapping = config.AUTO_RESULT_MAP.get(record.get("isXGSPONsupport", ""))
    if mapping:
        record.update(mapping)
    else:
        record.update({
            "Is2500Support": "NA",
            "defeatBuilding": "NA",
            "remark": f"接口无XGSPON标识",
        })
    return record


def build_error_record(keyword: str, message: str) -> Dict[str, Any]:
    """接口异常时的降级记录：不占用正常行键（floor/flat 留空），下次运行自动重试。"""
    return {
        "inputBuilding": keyword,
        "curBuilding": "",
        "curFloor": "",
        "curFlat": "",
        "Is2500Support": "",
        "defeatBuilding": "",
        "remark": f"错误:{message[:120]}",
    }


# ----------------------------------------------------------------------
# 单栋楼宇处理
# ----------------------------------------------------------------------
def process_building(
    client: CmhkApiClient,
    bot: BrowserBot,
    storage: ExcelStorage,
    keyword: str,
    judge_mode: bool = False,
) -> bool:
    """
    处理一栋楼宇：浏览器搜地址 → 后台 getAddressDetail → 采样 →
                浏览器内调用 getInstallInfo → 落库。

    :return: True=全部采样成功；False=存在失败（下次运行会重试该楼宇）
    """
    # ---- 1.【浏览器】打开页面 + 输入地址 + 点下拉第一项 ----
    try:
        bot.open_broadband_page()
        captured = bot.search_and_pick_first(keyword)
    except Exception as exc:
        msg = f"browser.search_and_pick_first: {exc}"
        print(f"    [浏览器自动化异常] {msg}")
        logger.exception("search_and_pick_first 失败")
        storage.upsert_row(build_error_record(keyword, msg))
        return False

    address_obj = captured.address_obj
    cur_building = captured.address_text or address_obj.get("value", "")
    if not address_obj:
        msg = f"未取到 searchAddress 响应（keyword={keyword}）"
        print(f"    [错误] {msg}")
        storage.upsert_row(build_error_record(keyword, msg))
        return False

    print(f"    curBuilding = {cur_building}")

    # ---- 1.5 地址匹配校验：系统返回地址必须包含输入楼宇名 ----
    # 模糊搜索在地址不存在时会返回近似候选，点第一项会"张冠李戴"
    # （例：输入"新科技广场"实际解析成"大有大廈"）。不匹配则降级记录。
    if not is_address_matched(keyword, cur_building):
        building_name = extract_building_name(keyword)
        remark = "eshop未查到该地址"
        print(f"    [降级] {remark}（输入[{building_name}]，"
              f"系统最近似返回[{cur_building}]）")
        logger.warning("地址未匹配(eshop未查到该地址): keyword=%s -> cur_building=%s",
                       keyword, cur_building)
        storage.upsert_row({
            "inputBuilding": keyword,
            "curBuilding": cur_building,
            "remark": remark,
        })
        return True

    client.throttle()

    # ---- 2.【后台】getAddressDetail：busiResp 保存为 busiRespObj ----
    try:
        detail_data = client.get_address_detail(address_obj)
    except ApiError as exc:
        msg = f"getAddressDetail: {exc}"
        print(f"    [接口异常] {msg}")
        storage.upsert_row(build_error_record(keyword, msg))
        return False
    busi_resp_obj = detail_data.get("busiResp") or {}

    # ---- 3. 采样 ----
    samples = pick_sampled_units(busi_resp_obj, seed=keyword)
    if not samples:
        print("    [提示] 该楼宇无楼层/单位数据")
        storage.upsert_row({
            "inputBuilding": keyword,
            "curBuilding": cur_building,
            "remark": "楼宇无楼层单位数据",
        })
        return True

    data = busi_resp_obj.get("busiDataResp") or {}
    print(f"    共 {len(data.get('homeUnitList', []))} 层，"
          f"采样 {len(samples)} 组: "
          + ", ".join(f"{f}层{p}室(组{g})" for f, p, g in samples))

    # ---- 4.【浏览器】等进入"选楼层/单位"步骤 ----
    try:
        bot.wait_for_floor_flat_step()
    except Exception as exc:
        msg = f"等楼层单位步骤超时: {exc}"
        print(f"    [浏览器状态异常] {msg}")
        # 即便页面没真正进入选楼层步骤，也先尝试一批 fetch，能跑就跑
        logger.warning("wait_for_floor_flat_step 超时，继续尝试 fetch")

    building_code = pick_building_code(address_obj.get("carrierInfo", []))
    all_ok = True

    # ---- 5+6. 每组采样：都重新走 search→select→(click_next|fetch) ----
    # 原因：点完「下一步」后页面跳到结果页，下一组的 el-select 不在，
    #      必须重新 search 才能进入输入步骤页。
    for idx, (floor, flat, group_label) in enumerate(samples, start=1):
        # 行级断点续跑：该采样行已写入则跳过
        if storage.has_row(keyword, floor, flat):
            print(f"    [{idx}/{len(samples)}] {floor}层{flat}室 已存在，跳过")
            continue

        bus_info = build_install_bus_info(address_obj, building_code, floor, flat)
        record: Dict[str, Any] = {
            "inputBuilding": keyword,
            "curBuilding": cur_building,
            "curFloor": floor,
            "curFlat": flat,
            "buildingCode": building_code,
            "ofcaCode": address_obj.get("ofcaCode", ""),
            "floorGroup": group_label,
            "checkedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }

        # 重新进入输入/选楼层步骤页：
        #   - 第 1 组：step 1 已 search+pick，step 4 已等过楼层步骤，直接复用当前页
        #   - 第 2 组起：结果页点「重新查詢」页内回退（省整页重载+观镜等待），
        #               失败再兜底 open_broadband_page()
        try:
            if idx == 1:
                # 已在楼层/单位步骤，确认一下即可；超时则兜底整页重载
                try:
                    bot.wait_for_floor_flat_step(timeout=20)
                except Exception:
                    bot.open_broadband_page()
                    bot.search_and_pick_first(keyword)
                    bot.wait_for_floor_flat_step()
            else:
                if not bot.back_to_search():
                    bot.open_broadband_page()
                bot.search_and_pick_first(keyword)
                bot.wait_for_floor_flat_step()
        except Exception as exc:
            logger.warning("重新进入输入步骤页失败: %s", exc)
            record["remark"] = f"错误:重新进入输入页:{str(exc)[:80]}"
            storage.upsert_row(record)
            all_ok = False
            continue

        # 视觉下拉选中
        try:
            bot.select_floor_flat_visually(floor=floor, flat=flat)
        except Exception as exc:
            logger.warning("select_floor_flat_visually 失败: %s", exc)
            record["remark"] = f"错误:下拉选择:{str(exc)[:80]}"
            storage.upsert_row(record)
            all_ok = False
            continue

        # ====== 两条分支：人工判定（默认）/ 全自动（--auto） ======
        if judge_mode:
            # ---- 人工判定模式：点「下一步」→ 跳结果页 → 注入判定面板 →
            #      等用户【终端输入】或【页面点击】，无输入则一直等 ----
            try:
                if not bot.click_next():
                    logger.warning("未找到「下一步」按钮（页面可能已跳转）")
            except Exception as exc:
                logger.warning("click_next 异常: %s", exc)
            try:
                bot.wait_for_result_step(timeout=15)
            except Exception as exc:
                logger.debug("等结果步骤超时（容忍）: %s", exc)

            # 注入浮动判定面板（4 按钮），失败仍可用终端输入
            bot.inject_judge_panel()

            prompt = (
                f"    [{idx}/{len(samples)}] {floor}层{flat}室(组{group_label}) → "
                f"判定 [Y]可卖2500M / [N]不可卖 / [E]异常 / [S]跳过"
                f"（终端输入，或在浏览器页面点按钮）: "
            )
            ans = wait_for_judgment(bot, prompt)

            if ans == "S":
                record["remark"] = "用户跳过"
                storage.upsert_row(record)
                continue
            record["Is2500Support"] = ans
            if ans == "N":
                record["defeatBuilding"] = "Y"
            record["remark"] = "人工判定"
            print(f"    [{idx}/{len(samples)}] {floor}层{flat}室(组{group_label}) "
                  f"Is2500Support={ans} (人工)")
            storage.upsert_row(record)
        else:
            # ---- 全自动模式：浏览器内 fetch getInstallInfo → 自动判定 ----
            try:
                install_data = bot.fetch_install_info_in_browser(bus_info)
            except Exception as exc:
                print(f"    [{idx}/{len(samples)}] {floor}层{flat}室 "
                      f"[getInstallInfo 失败] {exc}")
                record.update({"remark": f"错误:getInstallInfo:{str(exc)[:100]}"})
                storage.upsert_row(record)
                all_ok = False
                continue

            record.update(summarize_install(install_data))
            record = apply_auto_result(record)
            print(f"    [{idx}/{len(samples)}] {floor}层{flat}室(组{group_label}) "
                  f"isXGSPONsupport={record['isXGSPONsupport'] or '-'} "
                  f"-> {record['remark']}")
            storage.upsert_row(record)

        client.throttle()

    return all_ok


# ----------------------------------------------------------------------
# 主流程
# ----------------------------------------------------------------------
def run(args: argparse.Namespace) -> int:
    # ---- 初始化 ----
    lines = read_input_lines(args.input)
    if not lines:
        print(f"[错误] 输入文件为空或不存在：{args.input}")
        return 1
    if args.limit:
        lines = lines[: args.limit]
    total = len(lines)

    storage = ExcelStorage(path=args.output)
    progress = ProgressTracker(ignore_existing=args.redo)   # redo：忽略旧进度，全部重跑

    client = CmhkApiClient(
        search_token=args.token_search,
        detail_token=args.token_detail,
    )

    mode_label = ("人工判定（点下一步 + 终端输入/页面点击，无输入一直等）"
                  if args.judge else "全自动浏览器+后台混合")
    print(
        f"[启动] 输入 {total} 栋 | Excel: {args.output} | "
        f"已完成 {progress.completed_count()} 栋 | "
        f"模式: {mode_label} | "
        f"采样规则: 每 {config.FLOOR_GROUP_SIZE} 层随机 1 层 + 当层随机 1 单位"
    )

    bot: Optional[BrowserBot] = None
    try:
        # getInstallInfo 必须在浏览器内执行，因此浏览器始终启动
        bot = BrowserBot(
            install_token=args.token_install or config.XGIOG_TOKEN_INSTALL,
            headless=args.headless,
        )
        bot.start()

        for row_no, keyword in enumerate(lines, start=1):
            # ---- 楼宇级断点续跑 ----
            if not args.redo and progress.is_done(keyword):
                continue

            print(f"\n[{row_no}/{total}] inputBuilding = {keyword}")
            ok = False
            # ---- 浏览器自愈：失败时自动重启一次（覆盖用户误关 / Edge 崩溃 / 防御超时） ----
            for attempt in (1, 2):
                try:
                    ok = process_building(client, bot, storage, keyword,
                                          judge_mode=args.judge)
                    if ok:
                        break
                    # 失败后尝试浏览器自愈
                    if attempt == 1 and not bot.ensure_alive():
                        print("    [浏览器自愈失败，跳过该楼宇]")
                        break
                except KeyboardInterrupt:
                    raise
                except Exception as exc:                      # 兜底：单栋异常不终止全局
                    print(f"    [未预期异常] {exc}")
                    logger.exception("process_building 未预期异常")
                    storage.upsert_row(build_error_record(keyword, str(exc)))
                    if attempt == 1 and bot.ensure_alive():
                        continue
                    break

            if ok:
                progress.mark_done(keyword)               # 完成标记落盘
            time.sleep(config.REQUEST_INTERVAL)

    except KeyboardInterrupt:
        # Ctrl+C：Excel 每行已实时保存，这里只做优雅退出提示
        print("\n[中断] 检测到 Ctrl+C，进度已保存，下次运行将从断点继续")
    finally:
        if bot:
            bot.close()
    return 0


# ----------------------------------------------------------------------
# 命令行入口
# ----------------------------------------------------------------------
def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="中国移动香港家居宽带 2500M(XGS-PON) 可售情况批量排查工具",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("-i", "--input", default=config.INPUT_FILE,
                        help="楼宇名称输入文件（每行一个）")
    parser.add_argument("-o", "--output", default=config.EXCEL_PATH,
                        help="结果 Excel 输出路径")
    parser.add_argument("--token-search", default=None,
                        help="searchAddress 令牌 XGiOG2f705（过期时覆盖 config 默认值）")
    parser.add_argument("--token-detail", default=None,
                        help="getAddressDetail 令牌 XGiOG2f705（过期时覆盖 config 默认值）")
    parser.add_argument("--token-install", default=None,
                        help="getInstallInfo 令牌 XGiOG2f705（过期时覆盖 config 默认值）")
    parser.add_argument("--headless", action="store_true",
                        help="浏览器无头模式（注意：站点反爬可能拦截无头浏览器，慎用）")
    parser.add_argument("--limit", type=int, default=0,
                        help="只处理前 N 栋（调试用，0=全部）")
    parser.add_argument("--redo", action="store_true",
                        help="忽略已有结果，全部重新排查")
    parser.add_argument("--judge", action="store_true", default=True,
                        help="人工判定模式（默认）：每组采样后点「下一步」→ 跳到结果页 → "
                             "在终端输入 Y(可卖)/N(不可卖)/E(异常)/S(跳过)，"
                             "或在浏览器页面点注入的判定按钮；无输入则一直等待")
    parser.add_argument("--auto", action="store_true",
                        help="全自动模式：关闭人工判定，浏览器内 fetch getInstallInfo 自动判定")
    parser.add_argument("--logfile", default=None,
                        help="可选日志文件路径，默认只输出到 stdout")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    # --auto 关闭默认的人工判定模式
    if args.auto:
        args.judge = False
    # 配置 logging
    handlers = [logging.StreamHandler(sys.stdout)]
    if args.logfile:
        handlers.append(logging.FileHandler(args.logfile, encoding="utf-8"))
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        handlers=handlers,
        force=True,
    )
    return run(args)


if __name__ == "__main__":
    sys.exit(main())
