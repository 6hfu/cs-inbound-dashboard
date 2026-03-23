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
from datetime import date, datetime, timedelta
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


def _utc_range(start_date, end_date):
    """date → SOQL用UTC文字列に変換（JST→UTC -9h、end_dateは翌日0時まで）"""
    jst_start = datetime(start_date.year, start_date.month, start_date.day)
    jst_end = datetime(end_date.year, end_date.month, end_date.day) + timedelta(days=1)
    return (
        (jst_start - timedelta(hours=9)).strftime("%Y-%m-%dT%H:%M:%SZ"),
        (jst_end - timedelta(hours=9)).strftime("%Y-%m-%dT%H:%M:%SZ"),
    )


def _months_in_range(start_date, end_date):
    """日付範囲に含まれる(year, month)のリストを返す"""
    months = []
    d = start_date.replace(day=1)
    while d <= end_date:
        months.append((d.year, d.month))
        if d.month == 12:
            d = date(d.year + 1, 1, 1)
        else:
            d = date(d.year, d.month + 1, 1)
    return months


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

def _fetch_zoom_inbound(start_date, end_date):
    sf = get_sf()
    utc_start, utc_end = _utc_range(start_date, end_date)
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


def fetch_daily_call_rate(start_date, end_date):
    calls = _fetch_zoom_inbound(start_date, end_date)
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
    for dt_date in sorted(daily.keys()):
        d = daily[dt_date]
        dt = datetime.combine(dt_date, datetime.min.time())
        rate = (d["answered"] / d["total"] * 100) if d["total"] > 0 else 0
        rows.append({
            "日付": dt, "曜日": days_jp[dt.weekday()],
            "入電数": d["total"], "受電対応数": d["answered"],
            "受電率": round(rate, 2), "取りこぼし": d["total"] - d["answered"],
        })
    return pd.DataFrame(rows)


def fetch_hourly_call_rate(start_date, end_date):
    calls = _fetch_zoom_inbound(start_date, end_date)
    hourly = defaultdict(lambda: defaultdict(lambda: {"total": 0, "answered": 0}))

    for cid, info in calls.items():
        jst = _jst(info["ring_time"])
        if not jst or not (10 <= jst.hour < 19):
            continue
        hourly[jst.date()][jst.hour]["total"] += 1
        if "answered" in info["results"]:
            hourly[jst.date()][jst.hour]["answered"] += 1

    rows = []
    for dt_date in sorted(hourly.keys()):
        dt = datetime.combine(dt_date, datetime.min.time())
        for hr in sorted(hourly[dt_date].keys()):
            d = hourly[dt_date][hr]
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


def fetch_call_results_raw(start_date, end_date):
    """個別Taskレコードを日付付きで取得（集約前、キャッシュ用）"""
    sf = get_sf()
    utc_start, utc_end = _utc_range(start_date, end_date)

    desc = sf.Task.describe()
    juden_val = None
    result_labels = []
    for f in desc["fields"]:
        if f["name"] == TYPE_FIELD:
            for pv in f.get("picklistValues", []):
                if pv["active"] and pv["value"] in ("受電",):
                    juden_val = pv["value"]
                    break
            if not juden_val:
                active_vals = [pv["value"] for pv in f.get("picklistValues", []) if pv["active"]]
                if len(active_vals) >= 2:
                    juden_val = active_vals[1]
        if f["name"] == RESULT_FIELD:
            result_labels = [pv["value"] for pv in f.get("picklistValues", []) if pv["active"]]

    if not juden_val:
        return pd.DataFrame(), []

    result = sf.query_all(
        f"SELECT OwnerId, {RESULT_FIELD}, CreatedDate "
        f"FROM Task "
        f"WHERE Status = 'clok' "
        f"AND {TYPE_FIELD} = '{juden_val}' "
        f"AND CreatedDate >= {utc_start} AND CreatedDate < {utc_end} "
        f"AND {BIZ_HOUR_FILTER}"
    )

    owner_ids = list({r["OwnerId"] for r in result["records"]})
    names = _resolve_user_names(sf, owner_ids)

    rows = []
    for r in result["records"]:
        jst = _jst(r["CreatedDate"])
        rows.append({
            "担当者": names.get(r["OwnerId"], "不明"),
            "結果": r.get(RESULT_FIELD) or "未選択",
            "日付": jst.date() if jst else None,
        })

    return pd.DataFrame(rows), result_labels


def aggregate_call_results(raw_df, result_labels):
    """生データから担当者別に集約する"""
    if raw_df.empty:
        return pd.DataFrame(), result_labels

    grouped = raw_df.groupby(["担当者", "結果"]).size().reset_index(name="件数")

    owners = defaultdict(lambda: defaultdict(int))
    for _, r in grouped.iterrows():
        owners[r["担当者"]][r["結果"]] += r["件数"]

    rows = []
    for name, results in owners.items():
        total = sum(results.values())
        completed = results.get("完了", 0)
        if completed == 0:
            for k, v in results.items():
                if "完了" in k or (len(result_labels) > 7 and k == result_labels[7]):
                    completed = v
                    break

        comp_rate = (completed / total * 100) if total > 0 else 0
        row = {"担当者": name, "受電数": total, "完了": completed, "完了率": round(comp_rate, 2)}
        for k, v in results.items():
            if k != "完了":
                row[k] = v
        rows.append(row)

    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.fillna(0)
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

def _normalize_name(name):
    return name.replace("\u3000", " ").strip()


def fetch_shift_data(start_date, end_date):
    sf = get_sf()
    utc_start, utc_end = _utc_range(start_date, end_date)

    # 受電対応した人のUser IDを取得
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

    operator_name_set = {_normalize_name(n) for n in operator_names.values()}
    hr_map = {}
    for r in hr_result["records"]:
        if _normalize_name(r["Name"]) in operator_name_set:
            hr_map[r["Id"]] = r["Name"]

    if not hr_map:
        return pd.DataFrame()

    ids_str = "','".join(hr_map.keys())
    day_fields = ", ".join([f"Field{129 + i}__c" for i in range(31)])

    # 範囲に含まれる月ごとにシフトデータを取得
    staff_data = {}  # normalized_name -> dict
    daily_counts = defaultdict(int)  # date -> 出勤者数

    for yr, mo in _months_in_range(start_date, end_date):
        month_str = f"{mo}月"
        work_result = sf.query_all(
            f"SELECT Id, Name, Field128__c, Field160__c, Field161__c, "
            f"{day_fields} "
            f"FROM CustomObject11__c "
            f"WHERE Field2__c = '{month_str}' "
            f"AND Field128__c IN ('{ids_str}')"
        )

        for r in work_result["records"]:
            staff_id = r.get("Field128__c")
            staff_name = _normalize_name(hr_map.get(staff_id, r.get("Name", "不明")))

            if staff_name not in staff_data:
                staff_data[staff_name] = {
                    "担当者": staff_name,
                    "予定時間": 0,
                    "稼働日数": 0,
                    "実績時間": 0,
                    "日別稼働": {},
                }

            staff_data[staff_name]["予定時間"] += (r.get("Field160__c") or 0)

            for i in range(31):
                day_num = i + 1
                try:
                    d = date(yr, mo, day_num)
                except ValueError:
                    continue
                if d < start_date or d > end_date:
                    continue

                val = r.get(f"Field{129 + i}__c")
                if val:
                    try:
                        parts = val.replace(".000Z", "").split(":")
                        hours = int(parts[0]) + int(parts[1]) / 60
                        staff_data[staff_name]["日別稼働"][d] = round(hours, 2)
                        staff_data[staff_name]["実績時間"] += hours
                        staff_data[staff_name]["稼働日数"] += 1
                        daily_counts[d] += 1
                    except (ValueError, IndexError):
                        pass

    rows = []
    for data in staff_data.values():
        data["実績時間"] = round(data["実績時間"], 1)
        rows.append(data)

    df = pd.DataFrame(rows)
    if not df.empty:
        df.attrs["daily_staff_count"] = dict(daily_counts)
    return df


def fetch_future_shift_counts(start_date, end_date, group_members=None):
    """指定期間の日別出勤予定者数・稼働時間予定を取得

    group_members: グループに所属するメンバー名のセット（指定時はその人だけ集計）
    """
    sf = get_sf()

    hr_result = sf.query_all(
        "SELECT Id, Name FROM CustomObject10__c WHERE Field39__c = 'CS'"
    )
    hr_map = {r["Id"]: _normalize_name(r["Name"]) for r in hr_result["records"]}
    if group_members:
        hr_map = {k: v for k, v in hr_map.items() if v in group_members}
    if not hr_map:
        return {}

    ids_str = "','".join(hr_map.keys())
    day_fields = ", ".join([f"Field{129 + i}__c" for i in range(31)])

    daily_counts = defaultdict(int)
    daily_hours = defaultdict(float)

    for yr, mo in _months_in_range(start_date, end_date):
        month_str = f"{mo}月"
        work_result = sf.query_all(
            f"SELECT Field128__c, {day_fields} "
            f"FROM CustomObject11__c "
            f"WHERE Field2__c = '{month_str}' "
            f"AND Field128__c IN ('{ids_str}')"
        )

        for r in work_result["records"]:
            for i in range(31):
                day_num = i + 1
                try:
                    d = date(yr, mo, day_num)
                except ValueError:
                    continue
                if d < start_date or d > end_date:
                    continue
                val = r.get(f"Field{129 + i}__c")
                if val:
                    try:
                        parts = val.replace(".000Z", "").split(":")
                        hours = int(parts[0]) + int(parts[1]) / 60
                        daily_counts[d] += 1
                        daily_hours[d] += hours
                    except (ValueError, IndexError):
                        pass

    return {d: {"出勤予定数": daily_counts[d], "稼働時間予定": round(daily_hours[d], 1)}
            for d in sorted(daily_counts.keys())}


def fetch_shift_by_members(start_date, end_date, member_names):
    """指定メンバーのシフトデータを取得（Task実績不要、シフト表用）"""
    sf = get_sf()

    hr_result = sf.query_all(
        "SELECT Id, Name FROM CustomObject10__c WHERE Field39__c = 'CS'"
    )
    hr_map = {}
    for r in hr_result["records"]:
        name = _normalize_name(r["Name"])
        if name in member_names:
            hr_map[r["Id"]] = name
    if not hr_map:
        return pd.DataFrame()

    ids_str = "','".join(hr_map.keys())
    day_fields = ", ".join([f"Field{129 + i}__c" for i in range(31)])

    staff_data = {}
    daily_counts = defaultdict(int)

    for yr, mo in _months_in_range(start_date, end_date):
        month_str = f"{mo}月"
        work_result = sf.query_all(
            f"SELECT Field128__c, Field160__c, {day_fields} "
            f"FROM CustomObject11__c "
            f"WHERE Field2__c = '{month_str}' "
            f"AND Field128__c IN ('{ids_str}')"
        )

        for r in work_result["records"]:
            staff_id = r.get("Field128__c")
            staff_name = hr_map.get(staff_id, "不明")

            if staff_name not in staff_data:
                staff_data[staff_name] = {
                    "担当者": staff_name,
                    "予定時間": 0,
                    "稼働日数": 0,
                    "実績時間": 0,
                    "日別稼働": {},
                }

            staff_data[staff_name]["予定時間"] += (r.get("Field160__c") or 0)

            for i in range(31):
                day_num = i + 1
                try:
                    d = date(yr, mo, day_num)
                except ValueError:
                    continue
                if d < start_date or d > end_date:
                    continue

                val = r.get(f"Field{129 + i}__c")
                if val:
                    try:
                        parts = val.replace(".000Z", "").split(":")
                        hours = int(parts[0]) + int(parts[1]) / 60
                        staff_data[staff_name]["日別稼働"][d] = round(hours, 2)
                        staff_data[staff_name]["実績時間"] += hours
                        staff_data[staff_name]["稼働日数"] += 1
                        daily_counts[d] += 1
                    except (ValueError, IndexError):
                        pass

    rows = []
    for data in staff_data.values():
        data["実績時間"] = round(data["実績時間"], 1)
        rows.append(data)

    df = pd.DataFrame(rows)
    if not df.empty:
        df.attrs["daily_staff_count"] = dict(daily_counts)
    return df
