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


tab_rate, tab_results, tab_shift, tab_improve, tab_groups = st.tabs([
    "📈 受電率", "📊 コール処理実績", "🕐 稼働実績", "🎯 改善分析", "⚙️ グループ設定"
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
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

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
                    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

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
                    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

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
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
        with col2:
            fig = px.bar(group_summary, x="グループ", y="完了率", title="グループ別 完了率",
                         color="完了率", color_continuous_scale="Greens", text="完了率")
            fig.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
            fig.update_yaxes(range=[0, 100])
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

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
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
        with col2:
            fig = px.bar(
                filtered.sort_values("完了率", ascending=True),
                x="完了率", y="担当者", orientation="h",
                title="担当者別 完了率", color="グループ", text="完了率",
            )
            fig.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
            fig.update_layout(height=max(400, len(filtered) * 30))
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

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
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

        col1, col2 = st.columns(2)
        with col1:
            s = shift_df.sort_values("稼働日数", ascending=True)
            fig = px.bar(s, x="稼働日数", y="担当者", title="担当者別 稼働日数",
                         orientation="h", color="稼働日数", color_continuous_scale="Purples")
            fig.update_layout(height=max(400, len(s) * 25))
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
        with col2:
            s = shift_df.sort_values("実績時間", ascending=True)
            fig = go.Figure()
            fig.add_trace(go.Bar(y=s["担当者"], x=s["予定時間"], name="予定時間",
                                 orientation="h", marker_color="#BBDEFB"))
            fig.add_trace(go.Bar(y=s["担当者"], x=s["実績時間"], name="実績時間",
                                 orientation="h", marker_color="#1565C0"))
            fig.update_layout(title="予定時間 vs 実績時間", barmode="overlay",
                              height=max(400, len(s) * 25))
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

        st.subheader("稼働実績一覧")
        ds = shift_df[["担当者", "稼働日数", "予定時間", "実績時間"]].sort_values("実績時間", ascending=False)
        st.dataframe(ds, use_container_width=True, hide_index=True)


# ----- タブ4: 改善分析 -----
with tab_improve:
    st.subheader("受電率の改善ポイント")

    if hourly_df.empty or shift_df.empty or daily_df.empty:
        st.warning("データが不足しています（受電率・稼働実績の両方が必要です）")
    else:
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
            merged["日付_str"] = merged["日付"].dt.strftime("%m/%d") + "(" + merged["曜日"] + ")"

            # 危険度判定
            def calc_risk(row):
                if row["受電率"] < 80:
                    return "🔴 危険"
                elif row["受電率"] < 90:
                    return "🟡 注意"
                else:
                    return "🟢 良好"
            merged["判定"] = merged.apply(calc_risk, axis=1)

            # 危険日アラート
            danger_days = merged[merged["判定"].isin(["🔴 危険", "🟡 注意"])]
            if not danger_days.empty:
                st.error(f"受電率90%未満の日が **{len(danger_days)}日** あります")
                for _, row in danger_days.iterrows():
                    st.markdown(
                        f"- **{row['日付_str']}** 受電率 {row['受電率']:.1f}% / "
                        f"出勤{row['出勤者数']}名 / 取りこぼし{row['取りこぼし']}件"
                    )
            else:
                st.success("全日 受電率90%以上で問題なし")

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
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

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
