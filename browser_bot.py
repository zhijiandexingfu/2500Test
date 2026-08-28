# -*- coding: utf-8 -*-
"""
browser_bot —— Playwright 浏览器自动化模块。

职责（按用户在最新提示词中描述的流程）：
    ① 可视化打开 https://www.hk.chinamobile.com/tc/home-family/broadband
       并通过「观镜防御系统」挑战页（持久化上下文 + 反检测参数）
    ② 在 #searchInput 输入楼宇关键字
    ③ 等联想下拉弹出（.searchGoodsLine 列表，详见 config.SELECTORS）
    ④ 点击下拉第一项（同时浏览器自动触发 searchAddress 接口）
       —— 该接口响应通过 _on_response 监听器被自动捕获
    ⑤ 等页面跳转到"选楼层/单位"步骤（两个 .el-select 出现）
    ⑥ 在浏览器内通过 page.evaluate(fetch(...)) 调用 getInstallInfo
       —— 携带真实的浏览器 cookie、UA、omni-client-device-id，
          走「观镜防御系统」放行后的同源通道（专用令牌在 URL 上）

设计要点：
    * 单一持久化上下文（一栋楼共用一个 BrowserBot 实例）；
      关闭后下次启动复用 cookie，可跳过第二次的挑战页。
    * 所有 fetch/getInstallInfo 都在浏览器页面【同源】上下文内执行，
      严格按用户「该 URL 采用浏览器自动化测试的逻辑进行处理」。
    * 严谨的错误恢复：每次 evaluate 都包了重试循环（处理挑战页跳转
      销毁执行上下文的边界情况）。
"""

from __future__ import annotations

import json
import logging
import os
import time
from dataclasses import dataclass
from typing import Any, Dict, Optional, Tuple

from playwright.sync_api import (
    BrowserContext,
    Page,
    Response,
    TimeoutError as PWTimeoutError,
    sync_playwright,
)

import config


logger = logging.getLogger("xgps.browser")


# ============================================================================
# getInstallInfo 在浏览器内的 fetch 模板
#   说明：searchAddress 浏览器已经自动触发（点下拉第一项），
#   getAddressDetail 在主流程里走后台 requests（api_client）；
#   只有 getInstallInfo 是「按用户描述用浏览器自动化方式处理」。
# ============================================================================

_BROWSER_FETCH_JS = r"""
async (args) => {
    const resp = await fetch(args.url, {
        method: 'POST',
        credentials: 'include',        // 带上浏览器的 cookie（防御系统放行后才有 cookie）
        headers: args.headers,         // 由 Python 注入（accept/api-version/channelid/...）
        body: 'busInfo=' + encodeURIComponent(args.busInfo),
    });
    const text = await resp.text();
    try {
        const data = JSON.parse(text);
        return { ok: resp.ok, status: resp.status, data };
    } catch (e) {
        return { ok: resp.ok, status: resp.status, text: text.slice(0, 400) };
    }
}
"""


@dataclass
class CapturedAddress:
    """浏览器自动触发的 searchAddress 响应解析结果。"""
    address_text: str        # 第一项展示的完整地址字符串（与 input 可能略有不同）
    address_obj: Dict[str, Any]   # busiDataResp[0] 完整对象（用于 getAddressDetail）


class BrowserBot:
    """
    浏览器自动化机器人：一栋楼共用一个实例。
    """

    def __init__(
        self,
        install_token: Optional[str] = None,
        headless: bool = False,
        profile_dir: Optional[str] = None,
    ) -> None:
        self.install_token = install_token or config.XGIOG_TOKEN_INSTALL
        self.headless = headless
        self.profile_dir = profile_dir or config.BROWSER_PROFILE_DIR

        self._playwright = None
        self.context: Optional[BrowserContext] = None
        self.page: Optional[Page] = None

        # 浏览器捕获到的接口响应：{"search": {...}, "detail": {...}, "install": {...}}
        self.captured: Dict[str, Any] = {}

    # -------------------- 生命周期 --------------------

    def start(self) -> None:
        """启动持久化浏览器 + 反检测参数。"""
        if self.page is not None:
            return

        os.makedirs(self.profile_dir, exist_ok=True)

        # 删除可能残留的锁文件（Edge/Chrome 偶发）
        for lockfile in ("SingletonLock", "SingletonCookie", "SingletonSocket",
                         "lockfile"):
            p = os.path.join(self.profile_dir, lockfile)
            try:
                if os.path.exists(p):
                    os.remove(p)
            except OSError:
                pass

        self._playwright = sync_playwright().start()
        # 优先尝试 Edge；用户机器无 Edge/Edge 已占用时降级到 Chrome。
        for ch in ("msedge", "chrome"):
            try:
                self.context = self._playwright.chromium.launch_persistent_context(
                    user_data_dir=self.profile_dir,
                    channel=ch,
                    headless=self.headless,
                    args=["--disable-blink-features=AutomationControlled"],
                    ignore_default_args=["--enable-automation"],
                    locale="zh-HK",
                )
                logger.info("浏览器启动成功，channel=%s", ch)
                break
            except Exception as e:
                logger.warning("channel=%s 启动失败：%s", ch, e)
                continue
        else:
            raise RuntimeError("无法启动 msedge 或 chrome，请先关闭其他占用进程。")

        # 取主页面（持久化上下文可能自带一个 about:blank）
        self.page = (
            self.context.pages[0]
            if self.context.pages
            else self.context.new_page()
        )
        # 注册响应监听（捕获 searchAddress 等接口的真实响应）
        self.page.on("response", self._on_response)

    def is_alive(self) -> bool:
        """检查 context/page 是否还活着（未关闭）。"""
        if self.context is None or self.page is None or self._playwright is None:
            return False
        try:
            # page.context.is_closed() / page.is_closed() 是最权威的探活方式
            if self.page.is_closed():
                return False
            if self.context.browser is None:
                return False
            return True
        except Exception:
            return False

    def restart(self) -> None:
        """关闭并重启浏览器。常用于：用户误关 Edge / Edge 自动崩 / 防御系统超时掉线。"""
        logger.warning("浏览器不可用，尝试自动重启...")
        try:
            self.close()
        except Exception:
            pass
        # 给系统一点时间回收进程/锁
        time.sleep(2)
        self.start()
        logger.warning("浏览器已自动重启")

    def ensure_alive(self) -> bool:
        """如死亡则自动重启；返回 True 表示可用。"""
        if self.is_alive():
            return True
        try:
            self.restart()
            return self.is_alive()
        except Exception as exc:
            logger.error("浏览器自动重启失败：%s", exc)
            return False

    def close(self) -> None:
        """关闭浏览器 + Playwright。"""
        try:
            if self.context:
                self.context.close()
        except Exception:
            pass
        try:
            if self._playwright:
                self._playwright.stop()
        except Exception:
            pass
        self.context = None
        self.page = None
        self._playwright = None

    # -------------------- 内部：响应监听 --------------------

    def _on_response(self, response: Response) -> None:
        """捕获 searchAddress / getAddressDetail 响应。"""
        try:
            url = response.url or ""
        except Exception:
            return
        if "searchAddress" in url and response.status == 200:
            try:
                self.captured["search"] = response.json()
            except Exception:
                pass
        elif "getAddressDetail" in url and response.status == 200:
            try:
                self.captured["detail"] = response.json()
            except Exception:
                pass
        elif "getInstallInfo" in url and response.status == 200:
            try:
                self.captured["install"] = response.json()
            except Exception:
                pass

    # -------------------- ① 打开官网 / 等挑战通过 --------------------

    def open_broadband_page(self) -> None:
        """打开宽带页 + 等待「观镜防御系统」挑战页通过。"""
        page = self.page
        page.goto(config.BROADBAND_PAGE_URL, wait_until="domcontentloaded", timeout=90000)
        # 防御系统挑战可能持续 5~30s，需要等 URL 落到 broadband 且 readyState=complete
        for _ in range(60):
            try:
                title = page.title() or ""
                url = page.url
                if "观镜" not in title and "broadband" in url and \
                        page.evaluate("() => document.readyState") == "complete":
                    break
            except Exception:
                pass
            time.sleep(1)
        # 再额外等一会确保 SPA 水合完成、#searchInput 可交互
        page.wait_for_selector(config.SELECTORS["search_input"], timeout=30000)

    # -------------------- ② ③ ④ 输入地址 + 点下拉第一项 --------------------

    def search_and_pick_first(self, keyword: str) -> CapturedAddress:
        """
        输入楼宇关键字 → 等弹出联想下拉 → 点击【第一项】 →
            同时浏览器自动触发 searchAddress（在输入阶段）和
            getAddressDetail（在点击阶段），二者都被监听器捕获。

        注意页面真实流：
            输入关键词    -> 后端 searchAddress 自动请求（联想）
            点击下拉第一项 -> 后端 getAddressDetail 自动请求（携带该地址对象）

        优先使用浏览器捕获到的 searchAddress 响应作为地址对象；
        若 15 秒内都未捕获到，则回退到后台 requests 拉一次。
        """
        page = self.page
        # 先清空上栋楼的残留响应，避免误用
        self.captured.pop("search", None)
        self.captured.pop("detail", None)

        inp = page.locator(config.SELECTORS["search_input"])
        inp.click()
        # 清空 + 逐字输入，触发 input/change 事件让前端弹出搜索建议
        inp.fill("")
        inp.press_sequentially(keyword, delay=40)
        # 等联想下拉出现
        page.wait_for_selector(
            config.SELECTORS["dropdown_items"],
            state="visible",
            timeout=20000,
        )
        # 读取第一项文本后再点击
        first = page.locator(config.SELECTORS["dropdown_items"]).first
        address_text = (first.inner_text() or "").strip()
        # 点击第一项 → 浏览器自动触发 getAddressDetail
        first.click()
        # 同时也耐心等 searchAddress 响应（输入阶段触发）被收齐
        deadline = time.time() + 15
        while time.time() < deadline and "search" not in self.captured:
            time.sleep(0.1)

        # 解析 address 对象（优先用浏览器捕获到的 searchAddress 响应）
        address_obj: Dict[str, Any] = {}
        search_resp = self.captured.get("search")
        if isinstance(search_resp, dict):
            try:
                items = (search_resp.get("busiResp") or {}).get("busiDataResp") or []
                if items:
                    address_obj = items[0]
            except Exception:
                pass

        # 也接受「点击触发的 getAddressDetail」响应——它内部也含 carrierInfo/ofcaCode
        # 但通常 inputBuilding 文本已经足够定位，先用 searchAddress。
        if not address_obj:
            logger.warning("浏览器未捕获到 searchAddress，回退到后台 requests 拉取")
            from api_client import CmhkApiClient
            try:
                obj = CmhkApiClient(
                    search_token=config.XGIOG_TOKEN_SEARCH,
                ).first_address(keyword)
                if obj:
                    address_obj = obj
                    address_text = address_text or obj.get("value", keyword)
            except Exception as e:
                logger.warning("后台兜底拉 searchAddress 也失败：%s", e)

        return CapturedAddress(address_text=address_text or keyword, address_obj=address_obj)

    # -------------------- ⑤ 等页面进入"选楼层/单位"步骤 --------------------

    def wait_for_floor_flat_step(self, timeout: int = 30) -> None:
        """等到两个 .el-select（楼层 + 单位）渲染出来。"""
        page = self.page
        page.wait_for_function(
            "() => document.querySelectorAll('.el-select').length >= 2",
            timeout=timeout * 1000,
        )
        page.wait_for_timeout(500)  # 让动画结束

    def click_next(self, timeout: int = 15) -> bool:
        """
        点击页面的"下一步"按钮 (#btn_family_001)。
        关键：按钮存在 ≠ 可见 —— Vue 异步渲染 + CSS 过渡，需要先等它真正可见。
        返回 True 表示点击成功；False 表示按钮未出现（可能已跳到下一步/已自动跳转）。
        """
        page = self.page
        clicked = False
        for sel in (config.SELECTORS.get("next_button"),
                     config.SELECTORS.get("next_button_fallback")):
            try:
                btn = page.locator(sel)
                if btn.count() == 0:
                    continue
                # 先等按钮真正可见（offsetParent 非 null + 有尺寸）
                # 解决 select_floor_flat_visually 完成后 Vue 异步渲染导致按钮短暂 hidden 的问题
                try:
                    btn.first.wait_for(state="visible", timeout=timeout * 1000)
                except Exception:
                    logger.debug("next_button %s 等可见超时", sel)
                    continue
                # 滚到按钮位置（避免被 hero banner 挡住导致点不到）
                btn.first.scroll_into_view_if_needed()
                page.wait_for_timeout(200)
                btn.first.click()
                clicked = True
                logger.info("已点击「下一步」按钮: %s", sel)
                break
            except Exception as exc:
                logger.debug("next_button 探测 %s 失败: %s", sel, exc)
                continue
        if not clicked:
            return False
        # 给页面跳转/动画一点时间
        page.wait_for_timeout(800)
        return True

    def back_to_search(self, timeout: int = 15) -> bool:
        """
        在结果页点击「重新查詢」(#btn_family_003) 返回地址搜索步骤，
        省掉下一次 open_broadband_page() 的整页重载（含观镜等待）。
        成功回到搜索步骤返回 True；按钮不在/超时返回 False（调用方再兜底重载）。
        """
        page = self.page
        btn = page.locator(config.SELECTORS["requery_button"])
        if btn.count() == 0:
            return False
        try:
            btn.first.wait_for(state="visible", timeout=timeout * 1000)
            btn.first.click()
            # 回到搜索步骤：#searchInput 再次出现
            page.wait_for_selector(config.SELECTORS["search_input"],
                                   state="visible", timeout=timeout * 1000)
            page.wait_for_timeout(500)  # 让 SPA 水合
            return True
        except Exception as exc:
            logger.warning("back_to_search 失败（将兜底整页重载）: %s", exc)
            return False

    def wait_for_result_step(self, timeout: int = 30) -> None:
        """
        等页面进入"结果展示步骤"——按用户的描述，点完下一步后会跳到一个
        新的页面（呈现查询结果）。这里以页面 URL/标题变化作为信号。
        失败也不抛异常（不影响后续流程）。
        """
        page = self.page
        try:
            # 等下一步按钮消失（说明已离开前置步骤页面）
            page.wait_for_function(
                "() => !document.querySelector('#btn_family_001')",
                timeout=timeout * 1000,
            )
            page.wait_for_timeout(500)
        except Exception as exc:
            logger.debug("等结果步骤超时（可能页面结构有变）：%s", exc)

    # -------------------- ⑦ 结果页人工判定面板（注入） --------------------

    # 注入到结果页的浮动判定面板：4 个按钮，点击后写入 window.__xgps_judge
    _JUDGE_PANEL_JS = """
    () => {
        // 清空上一组的判定值
        window.__xgps_judge = null;
        // 移除旧面板（每组采样重新注入）
        const old = document.getElementById('xgps-judge-panel');
        if (old) old.remove();

        const div = document.createElement('div');
        div.id = 'xgps-judge-panel';
        div.style.cssText = (
            'position:fixed;right:20px;bottom:20px;z-index:2147483647;' +
            'background:#ffffff;border:2px solid #0085ba;border-radius:10px;' +
            'padding:14px 16px;box-shadow:0 4px 18px rgba(0,0,0,.35);' +
            'font-family:sans-serif;width:240px;'
        );
        const title = document.createElement('div');
        title.textContent = '判定：是否可卖 2500M？  (按 Y/N/E/S)';
        title.style.cssText = 'font-weight:bold;margin-bottom:10px;font-size:13px;color:#0085ba;';
        div.appendChild(title);

        const mk = (text, val, color) => {
            const b = document.createElement('button');
            b.id = 'xgps-judge-' + val;
            b.textContent = text;
            b.style.cssText = (
                'display:block;width:100%;margin:6px 0;padding:10px 0;border:0;' +
                'border-radius:6px;color:#fff;background:' + color + ';' +
                'font-size:14px;cursor:pointer;'
            );
            b.onmouseenter = () => { b.style.opacity = '0.85'; };
            b.onmouseleeleave = () => { b.style.opacity = '1'; };
            b.onclick = () => {
                window.__xgps_judge = val;
                // 点击后高亮，给用户视觉反馈
                div.querySelectorAll('button').forEach(x => x.style.outline = 'none');
                b.style.outline = '3px solid #0085ba';
                // 立即把焦点拿回 window 以便接着按 Y/N/E/S
                try { window.focus(); } catch (e) {}
            };
            div.appendChild(b);
        };
        mk('可卖 2500M (Y)', 'Y', '#00a651');
        mk('不可卖 2500M (N)', 'N', '#e23c3c');
        mk('其他情况待排 (E)', 'E', '#f39c12');
        mk('跳过 (S)', 'S', '#8a8a8a');

        document.body.appendChild(div);

        // 键盘监听：Y / N / E / S 直接触发判定（浏览器获焦时按键盘即可判定）
        // 兼容大小写 + 防抖（200ms 内重复按键只触发一次）
        if (!window.__xgps_keybound) {
            window.__xgps_keybound = true;
            window.__xgps_key_ts = 0;
            window.addEventListener('keydown', (e) => {
                // 忽略输入框/下拉里的按键，避免误触发
                const tag = (e.target && e.target.tagName || '').toUpperCase();
                if (tag === 'INPUT' || tag === 'TEXTAREA' || (e.target && e.target.isContentEditable)) {
                    return;
                }
                const k = (e.key || '').toUpperCase();
                if (k !== 'Y' && k !== 'N' && k !== 'E' && k !== 'S') return;
                const now = Date.now();
                if (now - window.__xgps_key_ts < 200) return;
                window.__xgps_key_ts = now;
                const btn = document.getElementById('xgps-judge-' + k);
                if (btn) { btn.click(); e.preventDefault(); }
            }, true);
        }
    }
    """

    def inject_judge_panel(self) -> None:
        """在结果页注入人工判定浮动面板（4 按钮：Y/N/E/S）。"""
        try:
            self.page.evaluate(self._JUDGE_PANEL_JS)
            logger.info("已注入人工判定面板（可卖/不可卖/待排/跳过）")
        except Exception as exc:
            # 注入失败不影响主流程（用户仍可在终端输入）
            logger.warning("注入判定面板失败（仍可用终端输入）: %s", exc)

    def read_judge_click(self) -> Optional[str]:
        """
        读取用户是否点击了页面判定按钮（main thread 轮询调用）。
        :return: 'Y'/'N'/'E'/'S' 之一；未点击返回 None。
        """
        try:
            return self.page.evaluate("() => window.__xgps_judge || null")
        except Exception:
            raise   # 异常向上抛，main 端的 wait_for_judgment 决定是否 fallback 到 stdin

    def reset_judge_click(self) -> None:
        """清空判定值（每组采样结束后调用，避免残留到下一组）。"""
        try:
            self.page.evaluate("() => { window.__xgps_judge = null; }")
        except Exception:
            pass

    # -------------------- ⑥ 浏览器内 fetch getInstallInfo --------------------

    def fetch_install_info_in_browser(
        self, bus_info: Dict[str, Any], retries: int = 5,
    ) -> Optional[Dict[str, Any]]:
        """
        在浏览器页面【同源】上下文内调用 getInstallInfo（按用户要求用浏览器方式）。
        带观镜防御系统放行后才会有的 cookie，规避风控。

        返回完整响应 JSON；连续失败抛 RuntimeError。
        """
        page = self.page
        url = (
            f"{config.BASE_URL}{config.API_GET_INSTALL_INFO}"
            f"?{config.TOKEN_PARAM}={self.install_token}"
        )
        # 把 fetch 需要注入的 headers 提前构造好（中文项目里保留 ASCII 单/双引号）
        browser_headers = dict(config.BROWSER_FETCH_HEADERS)
        js_args = {
            "url": url,
            "headers": browser_headers,
            "busInfo": json.dumps(bus_info, ensure_ascii=False),
        }

        last_err: Optional[Exception] = None
        for attempt in range(retries):
            try:
                result = page.evaluate(_BROWSER_FETCH_JS, js_args)
                if not isinstance(result, dict):
                    raise RuntimeError(f"非预期返回：{result!r}")
                if not result.get("ok"):
                    raise RuntimeError(
                        f"HTTP {result.get('status')}：{str(result.get('text') or result.get('data'))[:200]}"
                    )
                data = result.get("data")
                # 校验业务成功
                code = ((data or {}).get("busiResp") or {}).get("resCode", "")
                if code and code not in ("000000", "0", ""):
                    raise RuntimeError(f"业务失败 resCode={code}")
                return data
            except (PWTimeoutError, RuntimeError, Exception) as e:
                last_err = e
                logger.warning("浏览器内 fetch getInstallInfo 第 %d 次失败：%s", attempt + 1, e)
                time.sleep(1.5 + attempt * 0.5)  # 渐进退避
        raise RuntimeError(f"浏览器内 fetch getInstallInfo 失败 {retries} 次：{last_err}")

    # -------------------- 附加：选楼层/单位（视觉） --------------------

    def select_floor_flat_visually(self, floor: str, flat: str) -> None:
        """
        真实点击楼层 + 单位下拉（让用户在浏览器里看到"程序在操作"）。
        不会真去点击"下一步"——下一步触发的是网页跳转，
        我们用 fetch_install_info_in_browser 替代该请求，以节省时间。
        """
        page = self.page
        # 1) 楼层下拉（第 1 个）
        page.locator(config.SELECTORS["floor_select"]).scroll_into_view_if_needed()
        page.locator(config.SELECTORS["floor_select"]).click()
        page.wait_for_selector(
            config.SELECTORS["select_dropdown_item"], state="visible", timeout=10000,
        )
        page.locator(
            config.SELECTORS["select_dropdown_item"],
            has_text=floor,
        ).first.click()
        # 2) 单位下拉（第 2 个）
        page.wait_for_timeout(400)
        page.locator(config.SELECTORS["flat_select"]).scroll_into_view_if_needed()
        page.locator(config.SELECTORS["flat_select"]).click()
        page.wait_for_selector(
            config.SELECTORS["select_dropdown_item"], state="visible", timeout=10000,
        )
        page.locator(
            config.SELECTORS["select_dropdown_item"],
            has_text=flat,
        ).first.click()
        page.wait_for_timeout(400)
