"""Salesforce 実データ取得クライアント

データソース:
  - Task (Status='clok', Field3_del='受電'): コール処理実績（完了率など）
  - ZVC__Zoom_Call_Log__c: 受電率（Call_IDで一意着信を特定）
  - CustomObject11__c: 稼働実績
  - CustomObject10__c: 人事情報

営業時間: JST 10:00-19:00
"""

import os
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
from simple_salesforce import Salesforce

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass  # Streamlit Cloud では dotenv 不要

_sf = None
_user_name_cache = {}

# グループプリセット（アプリから変更可能）
DEFAULT_GROUPS = {
    "O既存": ["冨永 優", "村山 綾加", "森下 滉基", "益田 聖也"],
    "Z中堅": ["津越 花恋", "土門 真里菜", "大和田 祐一", "小林 美菜"],
    "Z新人": ["金子 愛実", "鎌田 未生", "長谷山 結唯", "國吉 空", "紺野 晴美",
              "小西 一希", "福澤 梨乃"],
    "O新人": ["柳原 良久", "小笠原 綾乃"],
    "バイトル": ["金澤 駿平", "結城 愛果", "csbt1", "csbt2", "csbt3"],
}

GROUPS_FILE = Path(__file__).parent / "groups.json"


def load_groups() -> dict:
    """グループ設定をファイルから読み込む。なければデフォルトを使用"""
    import json
    if GROUPS_FILE.exists():
        with open(GROUPS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return DEFAULT_GROUPS


def save_groups(groups: dict):
    """グループ設定をファイルに保存"""
    import json
    with open(GROUPS_FILE, "w", encoding="utf-8") as f:
        json.dump(groups, f, ensure_ascii=False, indent=2)


def _get_secret(key, default=None):
    """Streamlit Secrets → 環境変数 の順で取得"""
    try:
        import streamlit as st
        return st.secrets.get(key, os.getenv(key, default))
    except Exception:
        return os.getenv(key, default)


def get_sf():
    global _sf
    if _sf is None:
        _sf = Salesforce(
            username=_get_secret("SF_USERNAME"),
            password=_get_secret("SF_PASSWORD"),
            security_token=_get_secret("SF_SECURITY_TOKEN"),
            domain=_get_secret("SF_DOMAIN", "login"),
        )
    return _sf


def _jst(utc_str):
    if not utc_str:
        return None
    dt = datetime.strptime(utc_str[:19], "%Y-%m-%dT%H:%M:%S")
    return dt + timedelta(hours=9)


def _resolve_user_names(sf, user_ids):
    global _user_name_cache
    unknown = [uid for uid in user_ids if uid not in _user_name_cache]
    if unknown:
        for i in range(0, len(unknown), 100):
            batch = unknown[i:i+100]
            ids_str = "','".join(batch)
            result = sf.query(f"SELECT Id, Name FROM User WHERE Id IN ('{ids_str}')")
            for r in result["records"]:
                _user_name_cache[r["Id"]] = r["Name"]
    return {uid: _user_name_cache.get(uid, "不明") for uid in user_ids}


def _utc_range(year, month):
    jst_start = datetime(year, month, 1)
    jst_end = datetime(year + 1, 1, 1) if month == 12 else datetime(year, month + 1, 1)
    return (
        (jst_start - timedelta(hours=9)).strftime("%Y-%m-%dT%H:%M:%SZ"),
        (jst_end - timedelta(hours=9)).strftime("%Y-%m-%dT%H:%M:%SZ"),
    )


BIZ_HOUR_FILTER = "HOUR_IN_DAY(CreatedDate) >= 1 AND HOUR_IN_DAY(CreatedDate) < 10"


def fetch_all_cs_staff():
    """CS部署の全スタッフ名を返す（グループ設定用）"""
    sf = get_sf()
    result = sf.query_all(
        "SELECT Name FROM CustomObject10__c WHERE Field39__c = 'CS'"
    )
    names = set()
    for r in result["records"]:
        name = r["Name"].replace("\u3000", " ").strip()
        names.add(name)
    return sorted(names)


# =============================================
# 受電率（Zoom Call Log → Call_ID 重複除去）
# =============================================

def _fetch_zoom_inbound(year, month):
    sf = get_sf()
    utc_start, utc_end = _utc_range(year, month)
    result = sf.query_all(
        f"SELECT ZVC__Call_ID__c, ZVC__Call_Result__c, ZVC__Ring_Start_Time__c "
        f"FROM ZVC__Zoom_Call_Log__c "
        f"WHERE ZVC__Call_Type__c = 'inbound' "
        f"AND ZVC__Ring_Start_Time__c >= {utc_start} "
        f"AND ZVC__Ring_Start_Time__c < {utc_end}"
    )
    calls = {}
    for r in result["records"]:
        cid = r["ZVC__Call_ID__c"]
        res = r["ZVC__Call_Result__c"]
        ring = r.get("ZVC__Ring_Start_Time__c")
        if cid not in calls:
            calls[cid] = {"results": set(), "ring_time": ring}
        calls[cid]["results"].add(res)
        if ring and (calls[cid]["ring_time"] is None or ring < calls[cid]["ring_time"]):
            calls[cid]["ring_time"] = ring
    return calls


def fetch_daily_call_rate(year=2026, month=3):
    calls = _fetch_zoom_inbound(year, month)
    days_jp = ["月", "火", "水", "木", "金", "土", "日"]
    daily = defaultdict(lambda: {"total": 0, "answered": 0})

    for cid, info in calls.items():
        jst = _jst(info["ring_time"])
        if not jst or not (10 <= jst.hour < 19):
            continue
        key = jst.date()
        daily[key]["total"] += 1
        if "answered" in info["results"]:
            daily[key]["answered"] += 1

    rows = []
    for date in sorted(daily.keys()):
        d = daily[date]
        dt = datetime.combine(date, datetime.min.time())
        rate = (d["answered"] / d["total"] * 100) if d["total"] > 0 else 0
        rows.append({
            "日付": dt, "曜日": days_jp[dt.weekday()],
            "入電数": d["total"], "受電対応数": d["answered"],
            "受電率": round(rate, 2), "取りこぼし": d["total"] - d["answered"],
        })
    return pd.DataFrame(rows)


def fetch_hourly_call_rate(year=2026, month=3):
    calls = _fetch_zoom_inbound(year, month)
    hourly = defaultdict(lambda: defaultdict(lambda: {"total": 0, "answered": 0}))

    for cid, info in calls.items():
        jst = _jst(info["ring_time"])
        if not jst or not (10 <= jst.hour < 19):
            continue
        hourly[jst.date()][jst.hour]["total"] += 1
        if "answered" in info["results"]:
            hourly[jst.date()][jst.hour]["answered"] += 1

    rows = []
    for date in sorted(hourly.keys()):
        dt = datetime.combine(date, datetime.min.time())
        for hr in sorted(hourly[date].keys()):
            d = hourly[date][hr]
            rate = (d["answered"] / d["total"] * 100) if d["total"] > 0 else 0
            rows.append({
                "日付": dt, "時間帯": f"{hr}時", "時間帯_num": hr,
                "入電数": d["total"], "受電対応数": d["answered"],
                "受電率": round(rate, 2),
            })
    return pd.DataFrame(rows)


# =============================================
# コール処理実績（clok Task, Field3_del=受電）
# =============================================

# コール結果のフィールド値（Salesforceピックリスト準拠）
RESULT_FIELD = "Field4_del__c"
TYPE_FIELD = "Field3_del__c"


def fetch_call_results(year=2026, month=3):
    """担当者別のコール処理結果（完了、対応依頼、キャンセル等）"""
    sf = get_sf()
    utc_start, utc_end = _utc_range(year, month)

    # 受電のピックリスト値を取得
    desc = sf.Task.describe()
    juden_val = None
    result_labels = []
    for f in desc["fields"]:
        if f["name"] == TYPE_FIELD:
            for pv in f.get("picklistValues", []):
                if pv["active"] and pv["value"] in ("受電",):
                    juden_val = pv["value"]
                    break
                # Fallback: second value is 受電
            if not juden_val:
                active_vals = [pv["value"] for pv in f.get("picklistValues", []) if pv["active"]]
                if len(active_vals) >= 2:
                    juden_val = active_vals[1]  # 受電 is second
        if f["name"] == RESULT_FIELD:
            result_labels = [pv["value"] for pv in f.get("picklistValues", []) if pv["active"]]

    if not juden_val:
        return pd.DataFrame(), []

    result = sf.query_all(
        f"SELECT OwnerId, {RESULT_FIELD}, COUNT(Id) cnt "
        f"FROM Task "
        f"WHERE Status = 'clok' "
        f"AND {TYPE_FIELD} = '{juden_val}' "
        f"AND CreatedDate >= {utc_start} AND CreatedDate < {utc_end} "
        f"AND {BIZ_HOUR_FILTER} "
        f"GROUP BY OwnerId, {RESULT_FIELD}"
    )

    owners = defaultdict(lambda: defaultdict(int))
    for r in result["records"]:
        owners[r["OwnerId"]][r.get(RESULT_FIELD) or "未選択"] += r["cnt"]

    owner_ids = list(owners.keys())
    names = _resolve_user_names(sf, owner_ids)

    rows = []
    for uid, results in owners.items():
        name = names.get(uid, "不明")
        total = sum(results.values())
        completed = results.get("完了", 0)
        # 完了のピックリスト値が文字化けしている場合のフォールバック
        if completed == 0:
            for k, v in results.items():
                if "完了" in k or k == result_labels[7] if len(result_labels) > 7 else False:
                    completed = v
                    break

        comp_rate = (completed / total * 100) if total > 0 else 0

        row = {
            "担当者": name,
            "受電数": total,
            "完了": completed,
            "完了率": round(comp_rate, 2),
        }
        # 各結果を列に追加
        for k, v in results.items():
            if k != "完了":
                row[k] = v
        rows.append(row)

    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.fillna(0)
        # 数値列を整数に
        for col in df.columns:
            if col not in ("担当者", "完了率"):
                try:
                    df[col] = df[col].astype(int)
                except (ValueError, TypeError):
                    pass

    return df, result_labels


# =============================================
# 稼働実績
# =============================================

def fetch_shift_data(year=2026, month=3):
    sf = get_sf()

    # 受電対応した人のUser IDを取得
    utc_start, utc_end = _utc_range(year, month)
    desc = sf.Task.describe()
    juden_val = None
    for f in desc["fields"]:
        if f["name"] == TYPE_FIELD:
            active_vals = [pv["value"] for pv in f.get("picklistValues", []) if pv["active"]]
            if len(active_vals) >= 2:
                juden_val = active_vals[1]

    if not juden_val:
        return pd.DataFrame()

    op_result = sf.query_all(
        f"SELECT OwnerId FROM Task "
        f"WHERE Status = 'clok' AND {TYPE_FIELD} = '{juden_val}' "
        f"AND CreatedDate >= {utc_start} AND CreatedDate < {utc_end} "
        f"GROUP BY OwnerId"
    )
    operator_ids = {r["OwnerId"] for r in op_result["records"]}
    if not operator_ids:
        return pd.DataFrame()

    operator_names = _resolve_user_names(sf, list(operator_ids))

    hr_result = sf.query_all(
        "SELECT Id, Name FROM CustomObject10__c WHERE Field39__c = 'CS'"
    )

    def normalize(name):
        return name.replace("\u3000", " ").strip()

    operator_name_set = {normalize(n) for n in operator_names.values()}
    hr_map = {}
    for r in hr_result["records"]:
        if normalize(r["Name"]) in operator_name_set:
            hr_map[r["Id"]] = r["Name"]

    if not hr_map:
        return pd.DataFrame()

    ids_str = "','".join(hr_map.keys())
    month_str = f"{month}月"
    day_fields = ", ".join([f"Field{129 + i}__c" for i in range(31)])

    work_result = sf.query_all(
        f"SELECT Id, Name, Field128__c, Field164__c, Field160__c, Field161__c, "
        f"{day_fields} "
        f"FROM CustomObject11__c "
        f"WHERE Field2__c = '{month_str}' "
        f"AND Field128__c IN ('{ids_str}')"
    )

    rows = []
    daily_counts = defaultdict(int)
    for r in work_result["records"]:
        staff_id = r.get("Field128__c")
        staff_name = hr_map.get(staff_id, r.get("Name", "不明"))
        work_days = r.get("Field164__c") or 0
        scheduled = r.get("Field160__c") or 0
        actual = r.get("Field161__c") or 0

        daily_hours = {}
        for i in range(31):
            val = r.get(f"Field{129 + i}__c")
            if val:
                try:
                    parts = val.replace(".000Z", "").split(":")
                    hours = int(parts[0]) + int(parts[1]) / 60
                    daily_hours[i + 1] = round(hours, 2)
                    daily_counts[i + 1] += 1
                except (ValueError, IndexError):
                    pass

        rows.append({
            "担当者": staff_name,
            "稼働日数": work_days,
            "予定時間": scheduled,
            "実績時間": actual,
            "日別稼働": daily_hours,
        })

    df = pd.DataFrame(rows)
    if not df.empty:
        df.attrs["daily_staff_count"] = dict(daily_counts)
    return df
