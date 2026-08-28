# -*- coding: utf-8 -*-
"""
api_client —— 后台 requests 调用 searchAddress / getAddressDetail。

说明：
    * searchAddress 在主流程里由【浏览器】点击下拉第一项触发，本模块只作为兜底：
      当浏览器未能在规定时间内捕获到 searchAddress 响应时，回退到后台 requests 拉取。
    * getAddressDetail 全程使用后台 requests 调用，因为要批量获取所有 floor/flat
      （一栋楼一跑，避免每次采样都重建浏览器上下文）。

所有响应统一返回 JSON dict。失败抛出 ApiError 让上游感知。
"""

from __future__ import annotations

import json
import time
from typing import Any, Dict, Optional
from urllib.parse import urlencode

import requests

import config


class ApiError(Exception):
    """接口调用失败（含 HTTP 错误、JSON 解析错误、业务 resCode 异常）。"""


def _build_url(path: str, token: str) -> str:
    """拼装接口 URL：路径 + ?XGiOG2f705=token"""
    return f"{config.BASE_URL}{path}?{config.TOKEN_PARAM}={token}"


def _post_form(session: requests.Session, url: str, form_field: str, form_data: Dict) -> Dict[str, Any]:
    """
    统一的后台 POST 接口调用：
        POST <url>  Content-Type: application/x-www-form-urlencoded
        Body: <form_field>=<urlencoded(json)>
    """
    body = {form_field: json.dumps(form_data, ensure_ascii=False)}
    resp = session.post(
        url, data=body, timeout=config.REQUEST_TIMEOUT,
    )
    # 容错：即使 HTTP 200，也可能 resCode != 000000
    try:
        data = resp.json()
    except Exception as e:
        raise ApiError(f"接口返回非 JSON（HTTP {resp.status_code}）：{resp.text[:200]}") from e

    busi = data.get("busiResp") if isinstance(data, dict) else None
    if not busi:
        raise ApiError(f"接口响应缺少 busiResp：{json.dumps(data)[:200]}")
    code = busi.get("resCode") or busi.get("respCode")
    if code not in (None, "", "000000", "0"):
        raise ApiError(f"业务失败 resCode={code}：{busi.get('resMsg') or busi.get('respDesc') or ''}")
    return data


class CmhkApiClient:
    """中国移动香港宽带——后台接口客户端。"""

    def __init__(
        self,
        search_token: Optional[str] = None,
        detail_token: Optional[str] = None,
    ) -> None:
        self.search_token = search_token or config.XGIOG_TOKEN_SEARCH
        self.detail_token = detail_token or config.XGIOG_TOKEN_DETAIL
        # 复用 Session 维持 keep-alive、复用 cookie
        self.session = requests.Session()
        self.session.headers.update(config.DEFAULT_HEADERS)

    # -------------------- searchAddress --------------------
    def search_address(self, keyword: str) -> Dict[str, Any]:
        """
        调用 searchAddress 接口，返回完整响应 JSON。
        入参: {"keyword": "<楼宇关键字>"}
        """
        url = _build_url(config.API_SEARCH_ADDRESS, self.search_token)
        return _post_form(
            self.session, url, "busInfo", {"keyword": keyword}
        )

    def first_address(self, keyword: str) -> Optional[Dict[str, Any]]:
        """
        便捷方法：调 searchAddress 并取返回列表的第一个地址对象。
        无结果时返回 None。
        """
        data = self.search_address(keyword)
        busi = data.get("busiResp") or {}
        items = busi.get("busiDataResp") or []
        return items[0] if items else None

    # -------------------- getAddressDetail --------------------
    def get_address_detail(self, address_obj: Dict[str, Any]) -> Dict[str, Any]:
        """
        调 getAddressDetail，返回完整响应 JSON。
        address_obj 通常来自 searchAddress 返回的 busiDataResp[0]，至少含
            value / carrierInfo / ofcaCode 等字段。
        返回值的 `busiResp` 即用户提示词里说的 `busiRespObj`，含所有 floor/flat。
        """
        url = _build_url(config.API_GET_ADDRESS_DETAIL, self.detail_token)
        # 注入 clientType，保证与官网前端契约一致
        payload = dict(address_obj)
        payload.setdefault("clientType", config.CLIENT_TYPE)
        return _post_form(self.session, url, "orderStr", payload)

    # -------------------- 节流 --------------------
    def throttle(self, seconds: Optional[float] = None) -> None:
        time.sleep(seconds if seconds is not None else config.API_CALL_INTERVAL)
