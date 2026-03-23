"""CS入電 チーム分析ダッシュボード

起動: streamlit run app.py
営業時間: 10:00-19:00
"""

import json
from collections import defaultdict
from datetime import date, timedelta
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from salesforce_client import (
    fetch_daily_call_rate,
    fetch_hourly_call_rate,
    fetch_call_results_raw,
    aggregate_call_results,
    fetch_shift_data,
    fetch_future_shift_counts,
    fetch_all_cs_staff,
    load_groups,
    save_groups,
    DEFAULT_GROUPS,
)

st.set_page_config(page_title="CS入電 分析ダッシュボード", page_icon="📞", layout="wide")

st.title("📞 CS入電 分析ダッシュボード")

# --- 期間選択 ---
today = date.today()
this_month_start = today.replace(day=1)
if today.month == 1:
    last_month_start = date(today.year - 1, 12, 1)
else:
    last_month_start = date(today.year, today.month - 1, 1)
last_month_end = this_month_start - timedelta(days=1)
this_week_start = today - timedelta(days=today.weekday())
last_week_start = this_week_start - timedelta(days=7)
last_week_end = this_week_start - timedelta(days=1)

presets = {
    "今月": (this_month_start, today),
    "先月": (last_month_start, last_month_end),
    "今週": (this_week_start, today),
    "先週": (last_week_start, last_week_end),
    "直近7日": (today - timedelta(days=6), today),
    "直近14日": (today - timedelta(days=13), today),
    "カスタム": None,
}

col_preset, col_start, col_end, _ = st.columns([1.5, 1.2, 1.2, 3])
with col_preset:
    preset = st.selectbox("期間", list(presets.keys()))

if preset == "カスタム":
    with col_start:
        start_date = st.date_input("開始日", this_month_start)
    with col_end:
        end_date = st.date_input("終了日", today)
else:
    start_date, end_date = presets[preset]
    with col_start:
        st.date_input("開始日", start_date, disabled=True)
    with col_end:
        st.date_input("終了日", end_date, disabled=True)

col_caption, col_refresh = st.columns([5, 1])
with col_caption:
    st.caption(f"{start_date.strftime('%Y/%m/%d')} 〜 {end_date.strftime('%Y/%m/%d')} ｜ 営業時間 10:00-19:00 ｜ データソース: Salesforce")
with col_refresh:
    if st.button("🔄 最新データに更新"):
        st.cache_data.clear()
        st.rerun()


def _month_end(year, month):
    if month == 12:
        return date(year, 12, 31)
    return date(year, month + 1, 1) - timedelta(days=1)


def _months_covering(start_date, end_date):
    """日付範囲をカバーする(year, month)リスト"""
    months = []
    d = start_date.replace(day=1)
    while d <= end_date:
        months.append((d.year, d.month))
        if d.month == 12:
            d = date(d.year + 1, 1, 1)
        else:
            d = date(d.year, d.month + 1, 1)
    return months


@st.cache_data(ttl=21600)
def _load_month(year, month):
    """月単位でSalesforceからデータ取得（キャッシュ単位）"""
    s = date(year, month, 1)
    e = _month_end(year, month)
    daily = fetch_daily_call_rate(s, e)
    hourly = fetch_hourly_call_rate(s, e)
    results_raw, result_labels = fetch_call_results_raw(s, e)
    shift = fetch_shift_data(s, e)
    return daily, hourly, results_raw, result_labels, shift


def _filter_date(df, col, start_date, end_date):
    """DataFrameの日付列でフィルター"""
    if df.empty:
        return df
    mask = (df[col].dt.date >= start_date) & (df[col].dt.date <= end_date)
    return df[mask].reset_index(drop=True)


def _combine_and_filter_shifts(shifts, start_date, end_date):
    """複数月の稼働実績を統合し日付範囲でフィルター"""
    staff = {}
    daily_counts = defaultdict(int)

    for sdf in shifts:
        if sdf.empty:
            continue
        for _, row in sdf.iterrows():
            name = row["担当者"]
            if name not in staff:
                staff[name] = {"担当者": name, "予定時間": 0, "日別稼働": {}}
            staff[name]["予定時間"] += row["予定時間"]
            for d, h in row["日別稼働"].items():
                if start_date <= d <= end_date:
                    staff[name]["日別稼働"][d] = h

    rows = []
    for data in staff.values():
        hours = data["日別稼働"]
        if not hours:
            continue
        rows.append({
            "担当者": data["担当者"],
            "予定時間": data["予定時間"],
            "稼働日数": len(hours),
            "実績時間": round(sum(hours.values()), 1),
            "日別稼働": hours,
        })
        for d in hours:
            daily_counts[d] += 1

    df = pd.DataFrame(rows)
    if not df.empty:
        df.attrs["daily_staff_count"] = dict(daily_counts)
    return df


def load_data(start_date, end_date):
    """月単位キャッシュからデータを取得し、日付範囲でフィルター"""
    months = _months_covering(start_date, end_date)

    all_daily, all_hourly, all_raw, all_shifts = [], [], [], []
    result_labels = []

    for y, m in months:
        daily, hourly, raw, labels, shift = _load_month(y, m)
        all_daily.append(daily)
        all_hourly.append(hourly)
        all_raw.append(raw)
        all_shifts.append(shift)
        result_labels = labels

    # 受電率: concat + 日付フィルター
    daily_df = pd.concat(all_daily, ignore_index=True) if any(not d.empty for d in all_daily) else pd.DataFrame()
    hourly_df = pd.concat(all_hourly, ignore_index=True) if any(not h.empty for h in all_hourly) else pd.DataFrame()
    daily_df = _filter_date(daily_df, "日付", start_date, end_date)
    hourly_df = _filter_date(hourly_df, "日付", start_date, end_date)

    # コール処理: concat + 日付フィルター + 集約
    raw_df = pd.concat(all_raw, ignore_index=True) if any(not r.empty for r in all_raw) else pd.DataFrame()
    if not raw_df.empty:
        raw_df = raw_df[(raw_df["日付"] >= start_date) & (raw_df["日付"] <= end_date)]
    results_df, result_labels = aggregate_call_results(raw_df, result_labels)

    # 稼働: 統合 + 日付フィルター
    shift_df = _combine_and_filter_shifts(all_shifts, start_date, end_date)

    return daily_df, hourly_df, results_df, result_labels, shift_df


with st.spinner("Salesforceからデータ取得中..."):
    daily_df, hourly_df, results_df, result_labels, shift_df = load_data(start_date, end_date)


# グループ設定の読み込み
if "groups" not in st.session_state:
    st.session_state.groups = load_groups()


def get_group_for(name, groups):
    """担当者名からグループを返す"""
    for g, members in groups.items():
        if name in members:
            return g
    return "未割当"


tab_rate, tab_results, tab_shift, tab_shift_table, tab_improve, tab_groups = st.tabs([
    "📈 受電率", "📊 コール処理実績", "🕐 稼働実績", "📅 シフト表", "🎯 改善分析", "⚙️ グループ設定"
])


# ----- タブ1: 受電率 -----
with tab_rate:
    if daily_df.empty:
        st.warning("該当月のデータがありません")
    else:
        total_in = int(daily_df["入電数"].sum())
        total_ans = int(daily_df["受電対応数"].sum())
        total_missed = int(daily_df["取りこぼし"].sum())
        avg_rate = (total_ans / total_in * 100) if total_in > 0 else 0

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("受電率", f"{avg_rate:.1f}%")
        m2.metric("入電数", f"{total_in:,}")
        m3.metric("受電対応数", f"{total_ans:,}")
        m4.metric("取りこぼし", f"{total_missed:,}")

        # 日別 受電率 & 入電数
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=daily_df["日付"], y=daily_df["受電率"],
            name="受電率", mode="lines+markers",
            line=dict(color="#2196F3", width=2), yaxis="y1",
        ))
        fig.add_trace(go.Bar(
            x=daily_df["日付"], y=daily_df["入電数"],
            name="入電数", opacity=0.3, marker_color="#90CAF9", yaxis="y2",
        ))
        fig.update_layout(
            title="日別 受電率 & 入電数",
            yaxis=dict(title="受電率（%）", side="left"),
            yaxis2=dict(title="入電数", overlaying="y", side="right"),
            legend=dict(x=0, y=1.1, orientation="h"),
        )
        st.plotly_chart(fig, use_container_width=True, config={"staticPlot": True})

        # ヒートマップ & 時間帯別平均受電率
        col1, col2 = st.columns(2)
        with col1:
            if not hourly_df.empty:
                hour_labels = [f"{h}時" for h in range(10, 19)]
                filt = hourly_df[hourly_df["時間帯"].isin(hour_labels)]
                if not filt.empty:
                    pivot = filt.pivot_table(
                        index=filt["日付"].dt.strftime("%m/%d"),
                        columns="時間帯", values="受電率", aggfunc="mean",
                    )
                    ordered = [h for h in hour_labels if h in pivot.columns]
                    custom_scale = [
                        [0.0, "#d32f2f"],
                        [0.3, "#ff9800"],
                        [0.5, "#ffeb3b"],
                        [0.7, "#c8e6c9"],
                        [0.85, "#e8f5e9"],
                        [1.0, "#f1f8e9"],
                    ]
                    fig = px.imshow(pivot[ordered], title="時間帯×日付 受電率ヒートマップ",
                                   color_continuous_scale=custom_scale, aspect="auto",
                                   zmin=0, zmax=100,
                                   labels=dict(color="受電率(%)"))
                    st.plotly_chart(fig, use_container_width=True, config={"staticPlot": True})

        with col2:
            if not hourly_df.empty:
                hour_labels = [f"{h}時" for h in range(10, 19)]
                filt = hourly_df[hourly_df["時間帯"].isin(hour_labels)]
                if not filt.empty:
                    h_avg = filt.groupby("時間帯").agg(
                        平均受電率=("受電率", "mean"),
                    ).round(1).reindex(hour_labels).dropna().reset_index()
                    fig = px.bar(h_avg, x="時間帯", y="平均受電率", title="時間帯別 平均受電率",
                                 color="平均受電率", color_continuous_scale="RdYlGn", text="平均受電率")
                    fig.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
                    fig.update_yaxes(range=[0, 100])
                    st.plotly_chart(fig, use_container_width=True, config={"staticPlot": True})

        # 日別データ一覧
        st.subheader("日別データ一覧")
        dd = daily_df.copy()
        dd["日付"] = dd["日付"].dt.strftime("%m/%d")
        dd["受電率"] = dd["受電率"].apply(lambda x: f"{x:.1f}%")
        st.dataframe(dd[["日付", "曜日", "入電数", "受電対応数", "受電率", "取りこぼし"]],
                     use_container_width=True, hide_index=True)


# ----- タブ2: コール処理実績 -----
with tab_results:
    if results_df.empty:
        st.warning("該当月のデータがありません")
    else:
        groups = st.session_state.groups
        results_df["グループ"] = results_df["担当者"].apply(lambda x: get_group_for(x, groups))
        results_df = results_df[results_df["グループ"] != "未割当"]

        # 全体メトリクス
        total_calls = int(results_df["受電数"].sum())
        total_completed = int(results_df["完了"].sum())
        comp_rate = (total_completed / total_calls * 100) if total_calls > 0 else 0

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("受電数", f"{total_calls:,}")
        m2.metric("完了数", f"{total_completed:,}")
        m3.metric("完了率", f"{comp_rate:.1f}%")
        m4.metric("担当者数", f"{len(results_df)}名")

        # グループ別サマリー
        st.subheader("グループ別サマリー")
        display_cols = ["受電数", "完了"]
        optional_cols = ["対応依頼", "キャンセル受理", "キャンセル希望", "再コール", "処理のみ", "未選択"]
        for c in optional_cols:
            if c in results_df.columns:
                display_cols.append(c)

        group_summary = results_df.groupby("グループ")[display_cols].sum().reset_index()
        group_summary["完了率"] = (group_summary["完了"] / group_summary["受電数"] * 100).round(1)
        group_summary = group_summary.sort_values("受電数", ascending=False)

        col1, col2 = st.columns(2)
        with col1:
            fig = px.bar(group_summary, x="グループ", y="受電数", title="グループ別 受電数",
                         color="受電数", color_continuous_scale="Blues", text="受電数")
            fig.update_traces(textposition="outside")
            st.plotly_chart(fig, use_container_width=True, config={"staticPlot": True})
        with col2:
            fig = px.bar(group_summary, x="グループ", y="完了率", title="グループ別 完了率",
                         color="完了率", color_continuous_scale="Greens", text="完了率")
            fig.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
            fig.update_yaxes(range=[0, 100])
            st.plotly_chart(fig, use_container_width=True, config={"staticPlot": True})

        gs_display = group_summary.copy()
        gs_display["完了率"] = gs_display["完了率"].apply(lambda x: f"{x:.1f}%")
        st.dataframe(gs_display, use_container_width=True, hide_index=True)

        # 担当者別
        st.subheader("担当者別実績")
        group_filter = st.selectbox("グループで絞り込み", ["全て"] + list(groups.keys()))
        if group_filter != "全て":
            filtered = results_df[results_df["グループ"] == group_filter]
        else:
            filtered = results_df

        col1, col2 = st.columns(2)
        with col1:
            fig = px.bar(
                filtered.sort_values("受電数", ascending=True),
                x="受電数", y="担当者", orientation="h",
                title="担当者別 受電数", color="グループ", text="受電数",
            )
            fig.update_traces(textposition="outside")
            fig.update_layout(height=max(400, len(filtered) * 30))
            st.plotly_chart(fig, use_container_width=True, config={"staticPlot": True})
        with col2:
            fig = px.bar(
                filtered.sort_values("完了率", ascending=True),
                x="完了率", y="担当者", orientation="h",
                title="担当者別 完了率", color="グループ", text="完了率",
            )
            fig.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
            fig.update_layout(height=max(400, len(filtered) * 30))
            st.plotly_chart(fig, use_container_width=True, config={"staticPlot": True})

        # テーブル
        table_cols = ["グループ", "担当者", "受電数", "完了", "完了率"] + [
            c for c in optional_cols if c in filtered.columns
        ]
        table_df = filtered[table_cols].sort_values("受電数", ascending=False).copy()
        table_df["完了率"] = table_df["完了率"].apply(lambda x: f"{x:.1f}%")
        st.dataframe(table_df, use_container_width=True, hide_index=True)


# ----- タブ3: 稼働実績 -----
with tab_shift:
    if shift_df.empty:
        st.warning("該当月のデータがありません")
    else:
        m1, m2, m3 = st.columns(3)
        m1.metric("スタッフ数", f"{len(shift_df)}名")
        m2.metric("平均稼働日数", f"{shift_df['稼働日数'].mean():.1f}日")
        m3.metric("平均実績時間", f"{shift_df['実績時間'].mean():.1f}h")

        daily_counts = shift_df.attrs.get("daily_staff_count", {})
        if daily_counts:
            dc_df = pd.DataFrame([
                {"日付": d.strftime("%m/%d"), "出勤者数": c} for d, c in sorted(daily_counts.items())
            ])
            fig = px.bar(dc_df, x="日付", y="出勤者数", title="日別 出勤者数",
                         color="出勤者数", color_continuous_scale="Purples", text="出勤者数")
            fig.update_traces(textposition="outside")
            st.plotly_chart(fig, use_container_width=True, config={"staticPlot": True})

        col1, col2 = st.columns(2)
        with col1:
            s = shift_df.sort_values("稼働日数", ascending=True)
            fig = px.bar(s, x="稼働日数", y="担当者", title="担当者別 稼働日数",
                         orientation="h", color="稼働日数", color_continuous_scale="Purples")
            fig.update_layout(height=max(400, len(s) * 25))
            st.plotly_chart(fig, use_container_width=True, config={"staticPlot": True})
        with col2:
            s = shift_df.sort_values("実績時間", ascending=True)
            fig = go.Figure()
            fig.add_trace(go.Bar(y=s["担当者"], x=s["予定時間"], name="予定時間",
                                 orientation="h", marker_color="#BBDEFB"))
            fig.add_trace(go.Bar(y=s["担当者"], x=s["実績時間"], name="実績時間",
                                 orientation="h", marker_color="#1565C0"))
            fig.update_layout(title="予定時間 vs 実績時間", barmode="overlay",
                              height=max(400, len(s) * 25))
            st.plotly_chart(fig, use_container_width=True, config={"staticPlot": True})

        st.subheader("稼働実績一覧")
        ds = shift_df[["担当者", "稼働日数", "予定時間", "実績時間"]].sort_values("実績時間", ascending=False)
        st.dataframe(ds, use_container_width=True, hide_index=True)


# ----- タブ4: シフト表 -----
with tab_shift_table:
    if shift_df.empty:
        st.warning("該当期間のシフトデータがありません")
    else:
        groups = st.session_state.groups
        group_names = list(groups.keys())

        # グループ色の定義（淡め）
        GROUP_COLORS = {
            group_names[i] if i < len(group_names) else "": c
            for i, c in enumerate(["#5C9BD5", "#9B7FC4", "#6BAF6E", "#E08850", "#5BAAB5",
                                    "#D47070", "#7E6BAD", "#5DA06A"])
        }

        col_view, col_group, col_spacer = st.columns([1, 1.5, 3])
        with col_view:
            shift_view = st.radio("表示単位", ["月間", "週間"], horizontal=True, key="shift_view")
        with col_group:
            selected_group = st.selectbox("グループ", ["全体"] + group_names, key="shift_group")

        # 対象メンバーの絞り込み
        if selected_group == "全体":
            member_names = set()
            for members in groups.values():
                member_names.update(members)
        else:
            member_names = set(groups.get(selected_group, []))

        target_shifts = shift_df[shift_df["担当者"].isin(member_names)]

        if not target_shifts.empty:
            all_dates = set()
            for _, row in target_shifts.iterrows():
                all_dates.update(row["日別稼働"].keys())

            if all_dates:
                all_dates_sorted = sorted(all_dates)
                days_jp = ["月", "火", "水", "木", "金", "土", "日"]

                if shift_view == "週間":
                    weeks = sorted(set(d - timedelta(days=d.weekday()) for d in all_dates_sorted))
                    week_labels = [f"{w.strftime('%m/%d')}〜{(w + timedelta(days=6)).strftime('%m/%d')}" for w in weeks]
                    selected_week_idx = st.selectbox(
                        "週を選択", range(len(week_labels)),
                        format_func=lambda i: week_labels[i],
                        index=len(weeks) - 1, key="shift_week"
                    )
                    week_start = weeks[selected_week_idx]
                    display_dates = [week_start + timedelta(days=i) for i in range(7)
                                     if week_start + timedelta(days=i) in all_dates]
                else:
                    display_dates = all_dates_sorted

                if display_dates:
                    # グループ別にメンバーを整理
                    members_by_group = {}
                    for g in group_names:
                        g_members = set(groups[g])
                        g_shifts = target_shifts[target_shifts["担当者"].isin(g_members)]
                        if not g_shifts.empty:
                            members_by_group[g] = g_shifts.sort_values("担当者")

                    # HTMLテーブル構築
                    html = """
                    <style>
                    .shift-wrap { overflow: auto; max-height: 75vh; position: relative; }
                    .shift-table { border-collapse: separate; border-spacing: 0; width: max-content; font-size: 0.95em; }
                    .shift-table th { padding: 8px 6px; text-align: center; border: 1px solid #ddd;
                                      background: #f8f9fa; position: sticky; top: 0; z-index: 2; }
                    .shift-table th:first-child { position: sticky; left: 0; z-index: 3; background: #f8f9fa; }
                    .shift-table td { padding: 8px 6px; text-align: center; border: 1px solid #e0e0e0;
                                      font-size: 0.95em; }
                    .shift-table .name-cell { text-align: left; padding-left: 12px; white-space: nowrap;
                                               font-weight: 500; min-width: 100px; position: sticky;
                                               left: 0; z-index: 1; background: #fff;
                                               border-right: 2px solid #ccc; }
                    .shift-table .group-header { font-weight: bold; font-size: 0.95em; padding: 8px 12px;
                                                  text-align: left; color: white; position: sticky; left: 0; }
                    .shift-table .summary-row td { font-weight: bold; background: #f0f0f0; border-top: 2px solid #999; }
                    .shift-table .summary-row .name-cell { background: #f0f0f0; }
                    .shift-table .sat { background: #F0F6FF; }
                    .shift-table .sun { background: #FFF5F5; }
                    .shift-table .total-col { background: #FAFAFA; font-weight: 600; min-width: 50px; }
                    .cell-work { border-radius: 4px; padding: 3px 6px; display: inline-block; min-width: 32px; }
                    .cell-8h { background: #E8F5E9; color: #2E7D32; font-weight: 600; }
                    .cell-5h { background: #F3E8FF; color: #6B4D8A; }
                    .cell-short { background: #FFFDE7; color: #A68200; }
                    .cell-off { color: #d5d5d5; }
                    </style>
                    """

                    # ヘッダー行
                    html += '<div class="shift-wrap"><table class="shift-table">'
                    html += "<thead><tr><th style='min-width:90px;'>担当者</th>"
                    for d in display_dates:
                        dow = d.weekday()
                        dow_class = "sat" if dow == 5 else ("sun" if dow == 6 else "")
                        day_label = days_jp[dow]
                        html += f"<th class='{dow_class}'>{d.strftime('%m/%d')}<br><small>{day_label}</small></th>"
                    html += "<th class='total-col'>合計</th><th class='total-col'>日数</th></tr></thead>"
                    html += "<tbody>"

                    # グループごとにセクション表示
                    display_groups = [selected_group] if selected_group != "全体" else list(members_by_group.keys())

                    for g in display_groups:
                        if g not in members_by_group:
                            continue
                        g_shifts_df = members_by_group[g]
                        g_color = GROUP_COLORS.get(g, "#546E7A")

                        # グループヘッダー行
                        col_span = len(display_dates) + 3
                        html += f"<tr><td colspan='{col_span}' class='group-header' "
                        html += f"style='background:{g_color};'>{g}（{len(g_shifts_df)}名）</td></tr>"

                        for _, row in g_shifts_df.iterrows():
                            html += f"<tr><td class='name-cell'>{row['担当者']}</td>"
                            total_h = 0
                            work_days = 0
                            for d in display_dates:
                                h = row["日別稼働"].get(d, 0)
                                dow = d.weekday()
                                dow_class = "sat" if dow == 5 else ("sun" if dow == 6 else "")
                                if h > 0:
                                    total_h += h
                                    work_days += 1
                                    h_round = round(h, 1)
                                    if h >= 8:
                                        cell_class = "cell-work cell-8h"
                                    elif h >= 5:
                                        cell_class = "cell-work cell-5h"
                                    else:
                                        cell_class = "cell-work cell-short"
                                    html += f"<td class='{dow_class}'><span class='{cell_class}'>{h_round}</span></td>"
                                else:
                                    html += f"<td class='{dow_class} cell-off'>—</td>"
                            html += f"<td class='total-col'>{round(total_h, 1)}</td>"
                            html += f"<td class='total-col'>{work_days}</td></tr>"

                    # 集計行
                    html += "<tr class='summary-row'><td class='name-cell'>出勤者数</td>"
                    for d in display_dates:
                        cnt = sum(1 for _, row in target_shifts.iterrows() if row["日別稼働"].get(d, 0) > 0)
                        html += f"<td>{cnt}名</td>"
                    html += "<td></td><td></td></tr>"

                    html += "<tr class='summary-row'><td class='name-cell'>合計時間</td>"
                    for d in display_dates:
                        hrs = sum(row["日別稼働"].get(d, 0) for _, row in target_shifts.iterrows())
                        html += f"<td>{round(hrs, 1)}h</td>"
                    html += "<td></td><td></td></tr>"

                    html += "</tbody></table></div>"

                    # 凡例
                    html += """
                    <div style="margin-top:10px;display:flex;gap:16px;font-size:0.8em;color:#666;">
                        <span><span class="cell-work cell-8h" style="padding:2px 8px;border-radius:4px;">8.0</span> 8h以上</span>
                        <span><span class="cell-work cell-5h" style="padding:2px 8px;border-radius:4px;">6.0</span> 5〜8h</span>
                        <span><span class="cell-work cell-short" style="padding:2px 8px;border-radius:4px;">3.0</span> 5h未満</span>
                        <span style="color:#ccc;">— 休み</span>
                    </div>
                    """

                    st.markdown(html, unsafe_allow_html=True)

                    # グループ別集計サマリー
                    if selected_group == "全体" and len(members_by_group) > 1:
                        st.markdown("---")
                        st.subheader("グループ別集計")
                        group_summary_rows = []
                        for g in group_names:
                            if g not in members_by_group:
                                continue
                            g_shifts_df = members_by_group[g]
                            g_total_hours = g_shifts_df["実績時間"].sum()
                            g_avg_hours = g_shifts_df["実績時間"].mean()
                            g_avg_days = g_shifts_df["稼働日数"].mean()
                            g_count = len(g_shifts_df)

                            day_counts = []
                            for d in display_dates:
                                cnt = sum(1 for _, row in g_shifts_df.iterrows() if row["日別稼働"].get(d, 0) > 0)
                                day_counts.append(cnt)
                            avg_daily = sum(day_counts) / len(day_counts) if day_counts else 0

                            group_summary_rows.append({
                                "グループ": g,
                                "人数": f"{g_count}名",
                                "平均稼働日数": f"{g_avg_days:.1f}日",
                                "平均実績時間": f"{g_avg_hours:.1f}h",
                                "合計実績時間": f"{g_total_hours:.1f}h",
                                "日平均出勤": f"{avg_daily:.1f}名",
                            })
                        if group_summary_rows:
                            st.dataframe(pd.DataFrame(group_summary_rows),
                                         use_container_width=True, hide_index=True)
        else:
            st.info("選択したグループのシフトデータがありません")


# ----- タブ5: 改善分析 -----
with tab_improve:
    st.subheader("受電率の改善ポイント")

    if hourly_df.empty or shift_df.empty or daily_df.empty:
        st.warning("データが不足しています（受電率・稼働実績の両方が必要です）")
    else:
        # --- 0. 今後の注意日予測 ---
        st.subheader("今後の注意日予測")
        st.caption("過去の曜日別傾向とシフト予定から、受電率が下がりそうな日を予測します")

        days_jp = ["月", "火", "水", "木", "金", "土", "日"]

        # 曜日別平均（過去データから）
        weekday_stats = {}
        for dow in days_jp:
            wd = daily_df[daily_df["曜日"] == dow]
            if not wd.empty:
                weekday_stats[dow] = {
                    "平均入電数": wd["入電数"].mean(),
                    "平均受電率": wd["受電率"].mean(),
                }

        # 過去の稼働データから日別情報を取得
        past_daily_counts = shift_df.attrs.get("daily_staff_count", {})
        past_daily_hours = {}
        for _, row in shift_df.iterrows():
            for d, h in row["日別稼働"].items():
                past_daily_hours[d] = past_daily_hours.get(d, 0) + h

        # 過去の平均（出勤者数・稼働時間・一人あたり入電数）
        if past_daily_counts:
            avg_past_staff = sum(past_daily_counts.values()) / len(past_daily_counts)
            avg_past_hours = sum(past_daily_hours.values()) / len(past_daily_hours) if past_daily_hours else 0
        else:
            avg_past_staff = 0
            avg_past_hours = 0

        # 未来のシフト予定を取得（明日〜1ヶ月先）
        tomorrow = today + timedelta(days=1)
        future_end = date(today.year + (1 if today.month == 12 else 0),
                          1 if today.month == 12 else today.month + 1,
                          _month_end(today.year + (1 if today.month == 12 else 0),
                                     1 if today.month == 12 else today.month + 1).day)

        # グループ所属メンバーだけで集計
        groups = st.session_state.groups
        group_member_names = set()
        for members in groups.values():
            group_member_names.update(members)

        @st.cache_data(ttl=21600)
        def _load_future_shifts(s, e, members_tuple):
            return fetch_future_shift_counts(s, e, group_members=set(members_tuple))

        future_shifts = _load_future_shifts(tomorrow, future_end, tuple(sorted(group_member_names)))

        if future_shifts and weekday_stats and past_daily_counts:
            # 過去データから「一人あたり入電数 → 受電率」の関係を学習
            past_merged = daily_df[["日付", "曜日", "入電数", "受電率"]].copy()
            past_merged["日付_date"] = past_merged["日付"].dt.date
            past_dc = pd.DataFrame([{"日付_date": d, "出勤者数": c} for d, c in past_daily_counts.items()])
            past_merged = pd.merge(past_merged, past_dc, on="日付_date", how="inner")
            past_merged["一人あたり入電数"] = past_merged["入電数"] / past_merged["出勤者数"]

            # 一人あたり入電数の水準別に平均受電率を把握
            avg_per_person_all = past_merged["一人あたり入電数"].mean()

            def estimate_rate(per_person):
                """過去データの類似条件から受電率を推定"""
                if past_merged.empty:
                    return 90.0
                # 近い一人あたり入電数の日（±20%）の受電率平均
                similar = past_merged[
                    (past_merged["一人あたり入電数"] >= per_person * 0.8) &
                    (past_merged["一人あたり入電数"] <= per_person * 1.2)
                ]
                if len(similar) >= 3:
                    return similar["受電率"].mean()
                # データ不足時は曜日平均にフォールバック
                return past_merged["受電率"].mean()

            # 受電率90%以上の日の一人あたり入電数を目標値として算出
            good_days = past_merged[past_merged["受電率"] >= 90]
            if not good_days.empty:
                target_per_person = good_days["一人あたり入電数"].mean()
            else:
                target_per_person = avg_per_person_all * 0.85
            # 一人あたり平均稼働時間
            avg_hours_per_staff = avg_past_hours / avg_past_staff if avg_past_staff > 0 else 8.0

            forecast_rows = []
            for d, shift_info in future_shifts.items():
                dow = days_jp[d.weekday()]
                if dow not in weekday_stats:
                    continue

                expected_calls = weekday_stats[dow]["平均入電数"]
                staff_count = shift_info["出勤予定数"]
                hours_total = shift_info["稼働時間予定"]

                if staff_count == 0:
                    continue

                per_person = expected_calls / staff_count
                estimated_rate = estimate_rate(per_person)

                # 90%達成に必要な人数・時間を逆算
                needed_staff = int(expected_calls / target_per_person + 0.99)  # 切り上げ
                add_staff = max(0, needed_staff - staff_count)
                add_hours = round(add_staff * avg_hours_per_staff, 1)

                # 原因分析 + 追加目安
                risks = []
                if estimated_rate < 90:
                    if staff_count < avg_past_staff * 0.85:
                        risks.append(f"出勤者が少ない（{staff_count}名、過去平均{avg_past_staff:.0f}名）")
                    if per_person > avg_per_person_all * 1.15:
                        risks.append(f"一人あたり負荷が高い（{per_person:.1f}件、過去平均{avg_per_person_all:.1f}件）")
                    if expected_calls > daily_df["入電数"].mean() * 1.15:
                        risks.append(f"{dow}曜は入電が多い（予想{expected_calls:.0f}件）")
                    if not risks:
                        risks.append(f"入電数に対して人員が不足気味")

                # 判定
                if estimated_rate < 80:
                    risk_level = "🔴 要注意"
                elif estimated_rate < 90:
                    risk_level = "🟡 注意"
                else:
                    risk_level = "🟢 問題なし"

                forecast_rows.append({
                    "日付": d.strftime("%m/%d") + f"({dow})",
                    "予想入電数": f"{expected_calls:.0f}件",
                    "出勤予定": f"{staff_count}名",
                    "一人あたり": f"{per_person:.1f}件",
                    "予想受電率": f"{estimated_rate:.1f}%",
                    "追加人数目安": f"+{add_staff}名" if add_staff > 0 else "—",
                    "追加時間目安": f"+{add_hours}h" if add_hours > 0 else "—",
                    "判定": risk_level,
                    "risks": risks,
                    "_add_staff": add_staff,
                    "_add_hours": add_hours,
                })

            # 注意日
            alert_rows = [r for r in forecast_rows if r["判定"] != "🟢 問題なし"]
            critical_rows = [r for r in forecast_rows if r["判定"] == "🔴 要注意"]
            warn_rows = [r for r in forecast_rows if r["判定"] == "🟡 注意"]
            ok_rows = [r for r in forecast_rows if r["判定"] == "🟢 問題なし"]

            if alert_rows:
                # サマリー
                parts = []
                if critical_rows:
                    parts.append(f"🔴要注意 **{len(critical_rows)}日**")
                if warn_rows:
                    parts.append(f"🟡注意 **{len(warn_rows)}日**")
                parts.append(f"🟢問題なし {len(ok_rows)}日")
                st.warning("　".join(parts))

                # 注意日テーブル（追加目安付き）
                alert_df = pd.DataFrame(alert_rows)[
                    ["判定", "日付", "予想受電率", "予想入電数", "出勤予定", "一人あたり", "追加人数目安", "追加時間目安"]
                ]
                st.dataframe(alert_df, use_container_width=True, hide_index=True)

                # 特に危険な日（🔴）だけ詳細を表示
                if critical_rows:
                    st.markdown("**特に注意が必要な日：**")
                    for r in critical_rows:
                        st.markdown(
                            f"- **{r['日付']}** 予想受電率{r['予想受電率']}"
                            f"（{r['出勤予定']}で入電{r['予想入電数']}）"
                            f" → **{r['追加人数目安']}（{r['追加時間目安']}）追加で90%見込み**"
                        )
            else:
                st.success(f"今後のシフトに大きな問題は見られません（全{len(ok_rows)}日 問題なし）")
        else:
            st.info("今後のシフトデータが登録されていないか、過去データが不足しています")

        st.markdown("---")

        # --- 1. 時間帯別の優先度 ---
        st.subheader("時間帯別 増員の優先度")
        st.caption("受電率が低く入電が多い時間帯ほど、スタッフ増員の効果が高い")

        hour_labels = [f"{h}時" for h in range(10, 19)]
        filt = hourly_df[hourly_df["時間帯"].isin(hour_labels)]
        if not filt.empty:
            h_stats = filt.groupby("時間帯").agg(
                平均入電数=("入電数", "mean"),
                平均受電率=("受電率", "mean"),
                合計入電数=("入電数", "sum"),
                合計受電対応数=("受電対応数", "sum"),
            ).round(1).reindex(hour_labels).dropna().reset_index()
            h_stats["取りこぼし"] = h_stats["合計入電数"] - h_stats["合計受電対応数"]

            # 優先度テーブル
            priority = h_stats[["時間帯", "平均入電数", "平均受電率", "取りこぼし"]].copy()
            priority["優先度"] = priority.apply(
                lambda r: "🔴 高" if r["平均受電率"] < 80 and r["平均入電数"] > priority["平均入電数"].median()
                else ("🟡 中" if r["平均受電率"] < 90 else "🟢 低"), axis=1
            )
            priority["平均受電率"] = priority["平均受電率"].apply(lambda x: f"{x:.1f}%")
            st.dataframe(priority, use_container_width=True, hide_index=True)

        # --- 2. 危険日の検出 ---
        st.subheader("要注意日の検出")

        daily_counts = shift_df.attrs.get("daily_staff_count", {})
        if daily_counts:
            dr = daily_df[["日付", "曜日", "入電数", "受電対応数", "受電率", "取りこぼし"]].copy()
            dr["日付_date"] = dr["日付"].dt.date
            dc = pd.DataFrame([{"日付_date": d, "出勤者数": c} for d, c in daily_counts.items()])

            daily_hours = {}
            for _, row in shift_df.iterrows():
                for day, hours in row["日別稼働"].items():
                    daily_hours[day] = daily_hours.get(day, 0) + hours
            dh = pd.DataFrame([{"日付_date": d, "稼働時間合計": round(h, 1)} for d, h in daily_hours.items()])

            merged = pd.merge(dr, dc, on="日付_date", how="inner")
            merged = pd.merge(merged, dh, on="日付_date", how="left").fillna(0)
            merged["一人あたり入電数"] = (merged["入電数"] / merged["出勤者数"]).round(1)
            merged["一人あたり受電数"] = (merged["受電対応数"] / merged["出勤者数"]).round(1)
            merged["一時間あたり入電数"] = (merged["入電数"] / merged["稼働時間合計"].replace(0, float("nan"))).round(2)
            merged["日付_str"] = merged["日付"].dt.strftime("%m/%d") + "(" + merged["曜日"] + ")"

            # 平均値の計算
            avg_staff = merged["出勤者数"].mean()
            avg_hours = merged["稼働時間合計"].mean()
            avg_calls = merged["入電数"].mean()
            avg_per_person = merged["一人あたり受電数"].mean()
            avg_rate = merged["受電率"].mean()

            # 危険度判定
            def calc_risk(row):
                if row["受電率"] < 80:
                    return "🔴 危険"
                elif row["受電率"] < 90:
                    return "🟡 注意"
                else:
                    return "🟢 良好"
            merged["判定"] = merged.apply(calc_risk, axis=1)

            # 日別の原因分析
            def diagnose_day(row):
                causes = []
                if row["出勤者数"] < avg_staff * 0.85:
                    causes.append(f"出勤者不足（{row['出勤者数']:.0f}名、平均{avg_staff:.1f}名）")
                if row["稼働時間合計"] < avg_hours * 0.85:
                    causes.append(f"稼働時間不足（{row['稼働時間合計']:.1f}h、平均{avg_hours:.1f}h）")
                if row["入電数"] > avg_calls * 1.15:
                    causes.append(f"入電数が多い（{row['入電数']}件、平均{avg_calls:.0f}件）")
                if row["一人あたり受電数"] < avg_per_person * 0.8:
                    causes.append(f"一人あたり受電効率が低い（{row['一人あたり受電数']:.1f}件、平均{avg_per_person:.1f}件）")
                if not causes:
                    causes.append("複合的な要因（個別の突出した原因なし）")
                return causes
            merged["原因"] = merged.apply(diagnose_day, axis=1)

            # 危険日アラート
            danger_days = merged[merged["判定"].isin(["🔴 危険", "🟡 注意"])]
            if not danger_days.empty:
                danger_count = len(danger_days[danger_days["判定"] == "🔴 危険"])
                warn_count = len(danger_days[danger_days["判定"] == "🟡 注意"])
                total_missed_danger = int(danger_days["取りこぼし"].sum())
                parts = []
                if danger_count > 0:
                    parts.append(f"🔴 危険 **{danger_count}日**")
                if warn_count > 0:
                    parts.append(f"🟡 注意 **{warn_count}日**")
                st.error(f"受電率90%未満の日が **{len(danger_days)}日**（{' / '.join(parts)}）── 取りこぼし計 **{total_missed_danger}件**")

                # 要注意日カードを1列で表示
                cards_html = ""
                for _, row in danger_days.iterrows():
                    is_danger = row["判定"] == "🔴 危険"
                    border_color = "#EF5350" if is_danger else "#FFA726"
                    bg_color = "#FFF5F5" if is_danger else "#FFFBF0"
                    icon = "🔴" if is_danger else "🟡"
                    causes_text = "　".join(f"→ {c}" for c in row["原因"])
                    cards_html += f"""
                    <div style="border-left:4px solid {border_color};background:{bg_color};
                                padding:8px 14px;border-radius:6px;margin-bottom:6px;
                                display:flex;align-items:baseline;gap:12px;flex-wrap:wrap;">
                        <span style="font-weight:bold;font-size:1.0em;white-space:nowrap;">
                            {icon} {row['日付_str']}
                        </span>
                        <span style="font-size:0.9em;white-space:nowrap;">
                            受電率 <b>{row['受電率']:.1f}%</b>
                            ／出勤 <b>{row['出勤者数']:.0f}名</b>
                            ／取りこぼし <b>{row['取りこぼし']:.0f}件</b>
                        </span>
                        <span style="color:#555;font-size:0.83em;">{causes_text}</span>
                    </div>
                    """
                st.markdown(cards_html, unsafe_allow_html=True)
            else:
                st.success("全日 受電率90%以上で問題なし")

            # --- 期間全体のアドバイス ---
            st.subheader("この期間の総合アドバイス")
            total_missed = int(merged["取りこぼし"].sum())
            period_rate = (merged["受電対応数"].sum() / merged["入電数"].sum() * 100) if merged["入電数"].sum() > 0 else 0
            num_danger = len(danger_days)
            advices = []

            # 出勤者数の不足傾向
            low_staff_days = merged[merged["出勤者数"] < avg_staff * 0.85]
            if not low_staff_days.empty:
                advices.append(
                    f"📋 **出勤者数が平均を下回った日が{len(low_staff_days)}日**あります。"
                    f"シフト調整で出勤者数を平均{avg_staff:.0f}名以上に保てると受電率改善が見込めます。"
                )

            # 入電数の集中
            high_call_days = merged[merged["入電数"] > avg_calls * 1.15]
            if not high_call_days.empty:
                advices.append(
                    f"📞 **入電が集中した日が{len(high_call_days)}日**あります（平均{avg_calls:.0f}件超）。"
                    f"曜日や月初・月末の傾向を確認し、繁忙日に増員配置すると効果的です。"
                )

            # 一人あたり効率
            low_eff_days = merged[merged["一人あたり受電数"] < avg_per_person * 0.8]
            if not low_eff_days.empty and len(low_eff_days) >= 3:
                advices.append(
                    f"⚡ **一人あたり受電効率が低い日が{len(low_eff_days)}日**あります。"
                    f"離席時間の見直しや受電体制の改善で、一人あたり平均{avg_per_person:.1f}件を目指しましょう。"
                )

            # 稼働時間の不足
            low_hours_days = merged[merged["稼働時間合計"] < avg_hours * 0.85]
            if not low_hours_days.empty:
                advices.append(
                    f"🕐 **稼働時間が不足した日が{len(low_hours_days)}日**あります。"
                    f"短時間勤務の集中やシフト偏りがないか確認してください。"
                )

            # 全体が良好な場合
            if not advices:
                if period_rate >= 95:
                    advices.append("✅ 受電率95%以上を維持できています。現在の体制を継続してください。")
                elif period_rate >= 90:
                    advices.append("✅ 受電率90%以上で安定しています。あと少しの改善で95%到達も狙えます。")

            # 総括
            if num_danger > 0:
                summary = (
                    f"期間内の受電率は **{period_rate:.1f}%**、取りこぼしは合計 **{total_missed}件** です。"
                    f"要注意日{num_danger}日の主な原因に対処することで改善が見込めます。"
                )
            else:
                summary = f"期間内の受電率は **{period_rate:.1f}%**、取りこぼしは合計 **{total_missed}件** です。"

            st.info(summary)
            for advice in advices:
                st.markdown(advice)

            # メインチャート
            fig = go.Figure()
            colors = ["#EF5350" if j != "🟢 良好" else "#CE93D8" for j in merged["判定"]]
            fig.add_trace(go.Bar(
                x=merged["日付_str"], y=merged["稼働時間合計"],
                name="稼働時間(h)", marker_color=colors, opacity=0.6, yaxis="y2",
            ))
            fig.add_trace(go.Scatter(
                x=merged["日付_str"], y=merged["受電率"],
                name="受電率(%)", mode="lines+markers",
                line=dict(color="#1565C0", width=3), yaxis="y1",
            ))
            fig.add_hline(y=90, line_dash="dash", line_color="green", opacity=0.5,
                          annotation_text="目標90%", yref="y1")
            fig.update_layout(
                title="日別 稼働時間 × 受電率（赤 = 要注意日）",
                yaxis=dict(title="受電率（%）", side="left", range=[0, 100]),
                yaxis2=dict(title="稼働時間(h)", overlaying="y", side="right"),
                legend=dict(x=0, y=1.15, orientation="h"),
                height=450, xaxis=dict(tickangle=45),
            )
            st.plotly_chart(fig, use_container_width=True, config={"staticPlot": True})

            # 受電率90%以上 vs 未満の比較
            good = merged[merged["受電率"] >= 90]
            bad = merged[merged["受電率"] < 90]
            if not good.empty and not bad.empty:
                st.subheader("受電率 90%以上 と 未満の日の比較")
                comp = pd.DataFrame({
                    "区分": ["90%以上の日", "90%未満の日"],
                    "日数": [len(good), len(bad)],
                    "平均出勤者数": [f"{good['出勤者数'].mean():.1f}名", f"{bad['出勤者数'].mean():.1f}名"],
                    "平均稼働時間": [f"{good['稼働時間合計'].mean():.1f}h", f"{bad['稼働時間合計'].mean():.1f}h"],
                    "平均入電数": [f"{good['入電数'].mean():.0f}件", f"{bad['入電数'].mean():.0f}件"],
                    "一人あたり入電数": [f"{good['一人あたり入電数'].mean():.1f}件", f"{bad['一人あたり入電数'].mean():.1f}件"],
                })
                st.dataframe(comp, use_container_width=True, hide_index=True)

            # 日別一覧
            table = merged[["判定", "日付_str", "入電数", "受電対応数", "取りこぼし", "受電率",
                            "出勤者数", "稼働時間合計", "一人あたり入電数"]].copy()
            table.columns = ["判定", "日付", "入電数", "受電対応数", "取りこぼし", "受電率(%)",
                             "出勤者数", "稼働時間(h)", "一人あたり入電"]
            table["受電率(%)"] = table["受電率(%)"].apply(lambda x: f"{x:.1f}%")
            table = table.sort_values("日付")
            st.dataframe(table, use_container_width=True, hide_index=True)


# ----- タブ5: グループ設定 -----
with tab_groups:
    st.subheader("グループ設定")
    st.caption("担当者のグループ割り当てを変更できます。変更は保存ボタンで反映されます。")

    groups = st.session_state.groups

    # CS全スタッフ + コール実績の担当者を統合
    all_staff = set()
    try:
        all_staff.update(fetch_all_cs_staff())
    except Exception:
        pass
    if not results_df.empty:
        all_staff.update(results_df["担当者"].tolist())

    assigned = set()
    for members in groups.values():
        assigned.update(members)

    unassigned = all_staff - assigned
    if unassigned:
        st.warning(f"未割当の担当者: {', '.join(sorted(unassigned))}")

    # グループ並び替え
    st.markdown("**グループの表示順**")
    group_names = list(groups.keys())
    reordered_names = group_names.copy()

    cols_order = st.columns([4, 1, 1])
    with cols_order[0]:
        move_target = st.selectbox("移動するグループ", group_names, key="move_target")
    with cols_order[1]:
        if st.button("🔼 上へ"):
            idx = reordered_names.index(move_target)
            if idx > 0:
                reordered_names[idx], reordered_names[idx - 1] = reordered_names[idx - 1], reordered_names[idx]
                new_order = {k: groups[k] for k in reordered_names}
                save_groups(new_order)
                st.session_state.groups = new_order
                st.rerun()
    with cols_order[2]:
        if st.button("🔽 下へ"):
            idx = reordered_names.index(move_target)
            if idx < len(reordered_names) - 1:
                reordered_names[idx], reordered_names[idx + 1] = reordered_names[idx + 1], reordered_names[idx]
                new_order = {k: groups[k] for k in reordered_names}
                save_groups(new_order)
                st.session_state.groups = new_order
                st.rerun()

    st.caption(f"現在の順序: {' → '.join(group_names)}")
    st.markdown("---")

    # グループごとの編集
    new_groups = {}
    for group_name in group_names:
        members = groups[group_name]
        st.markdown(f"**{group_name}**")
        available = sorted(all_staff | set(members))
        selected = st.multiselect(
            f"{group_name} のメンバー",
            options=available,
            default=members,
            key=f"group_{group_name}",
        )
        new_groups[group_name] = selected

    # 新規グループ追加
    st.markdown("---")
    st.markdown("**新しいグループを追加**")
    new_name = st.text_input("グループ名", key="new_group_name")
    if new_name and new_name not in new_groups:
        new_members = st.multiselect(
            f"{new_name} のメンバー",
            options=sorted(all_staff),
            key=f"group_new_{new_name}",
        )
        if new_members:
            new_groups[new_name] = new_members

    # 保存
    col1, col2 = st.columns(2)
    with col1:
        if st.button("💾 保存", type="primary"):
            new_groups = {k: v for k, v in new_groups.items() if v}
            save_groups(new_groups)
            st.session_state.groups = new_groups
            st.success("保存しました")
            st.rerun()
    with col2:
        if st.button("🔄 デフォルトに戻す"):
            save_groups(DEFAULT_GROUPS)
            st.session_state.groups = DEFAULT_GROUPS
            st.success("デフォルトに戻しました")
            st.rerun()

    # 現在の設定表示
    st.markdown("---")
    st.subheader("現在の設定")
    for g, members in groups.items():
        st.markdown(f"**{g}**: {', '.join(members)}")
