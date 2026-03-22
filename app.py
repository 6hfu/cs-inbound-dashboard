"""CS入電 チーム分析ダッシュボード

起動: streamlit run app.py
営業時間: 10:00-19:00
"""

import json
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from salesforce_client import (
    fetch_daily_call_rate,
    fetch_hourly_call_rate,
    fetch_call_results,
    fetch_shift_data,
    fetch_all_cs_staff,
    load_groups,
    save_groups,
    DEFAULT_GROUPS,
)

st.set_page_config(page_title="CS入電 分析ダッシュボード", page_icon="📞", layout="wide")

st.title("📞 CS入電 分析ダッシュボード")

col_y, col_m, _ = st.columns([1, 1, 4])
with col_y:
    year = st.selectbox("年", [2026, 2025], index=0)
with col_m:
    month = st.selectbox("月", list(range(1, 13)), index=2)

st.caption(f"{year}年{month}月 ｜ 営業時間 10:00-19:00 ｜ データソース: Salesforce 活動の記録")


@st.cache_data(ttl=1800)
def load_data(y, m):
    daily_df = fetch_daily_call_rate(y, m)
    hourly_df = fetch_hourly_call_rate(y, m)
    results_df, result_labels = fetch_call_results(y, m)
    shift_df = fetch_shift_data(y, m)
    return daily_df, hourly_df, results_df, result_labels, shift_df

with st.spinner("Salesforceからデータ取得中..."):
    daily_df, hourly_df, results_df, result_labels, shift_df = load_data(year, month)


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
    st.subheader("受電率データ（営業時間 10:00-19:00）")

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
        st.plotly_chart(fig, use_container_width=True)

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
                    # 高い=薄い緑白、低い=赤 のカラースケール
                    custom_scale = [
                        [0.0, "#d32f2f"],   # 0% 赤
                        [0.3, "#ff9800"],   # 30% オレンジ
                        [0.5, "#ffeb3b"],   # 50% 黄
                        [0.7, "#c8e6c9"],   # 70% 薄い緑
                        [0.85, "#e8f5e9"],  # 85% とても薄い緑
                        [1.0, "#f1f8e9"],   # 100% ほぼ白
                    ]
                    fig = px.imshow(pivot[ordered], title="時間帯×日付 受電率ヒートマップ",
                                   color_continuous_scale=custom_scale, aspect="auto",
                                   zmin=0, zmax=100,
                                   labels=dict(color="受電率(%)"))
                    st.plotly_chart(fig, use_container_width=True)

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
                    st.plotly_chart(fig, use_container_width=True)

        # 時間帯別サマリー
        if not hourly_df.empty:
            st.subheader("時間帯別サマリー")
            hour_labels = [f"{h}時" for h in range(10, 19)]
            filt = hourly_df[hourly_df["時間帯"].isin(hour_labels)]
            if not filt.empty:
                hs = filt.groupby("時間帯").agg(
                    入電数=("入電数", "sum"), 受電対応数=("受電対応数", "sum"),
                ).reindex(hour_labels).dropna().reset_index()
                hs["受電率"] = (hs["受電対応数"] / hs["入電数"] * 100).round(1).apply(lambda x: f"{x:.1f}%")
                hs["シェア"] = (hs["入電数"] / hs["入電数"].sum() * 100).round(2).apply(lambda x: f"{x:.2f}%")
                st.dataframe(hs, use_container_width=True, hide_index=True)

        # 曜日別
        st.subheader("曜日別分析")
        wo = ["月", "火", "水", "木", "金", "土", "日"]
        wd = daily_df.groupby("曜日").agg(平均受電率=("受電率", "mean"), 平均入電数=("入電数", "mean")).round(1)
        wd = wd.reindex(wo).dropna().reset_index()
        if not wd.empty:
            fig = px.bar(wd, x="曜日", y="平均受電率", title="曜日別 平均受電率",
                         text="平均受電率", color="平均受電率", color_continuous_scale="RdYlGn")
            fig.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
            fig.update_yaxes(range=[0, 100])
            st.plotly_chart(fig, use_container_width=True)

        # 日別テーブル
        st.subheader("日別データ一覧")
        dd = daily_df.copy()
        dd["日付"] = dd["日付"].dt.strftime("%m/%d")
        dd["受電率"] = dd["受電率"].apply(lambda x: f"{x:.2f}%")
        st.dataframe(dd[["日付", "曜日", "入電数", "受電対応数", "受電率", "取りこぼし"]],
                     use_container_width=True, hide_index=True)


# ----- タブ2: コール処理実績 -----
with tab_results:
    st.subheader("コール処理実績（営業時間内）")

    if results_df.empty:
        st.warning("該当月のデータがありません")
    else:
        groups = st.session_state.groups

        # グループ列を追加
        results_df["グループ"] = results_df["担当者"].apply(lambda x: get_group_for(x, groups))

        # 全体メトリクス
        total_calls = int(results_df["受電数"].sum())
        total_completed = int(results_df["完了"].sum())
        comp_rate = (total_completed / total_calls * 100) if total_calls > 0 else 0

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("受電数", f"{total_calls:,}")
        m2.metric("完了数", f"{total_completed:,}")
        m3.metric("完了率", f"{comp_rate:.1f}%")
        m4.metric("担当者数", f"{len(results_df)}名")

        # --- グループ別サマリー ---
        st.subheader("グループ別サマリー")
        display_cols = ["受電数", "完了"]
        # 対応依頼等のカラムが存在すれば追加
        optional_cols = ["対応依頼", "キャンセル受理", "キャンセル希望", "再コール", "処理のみ", "未選択"]
        for c in optional_cols:
            if c in results_df.columns:
                display_cols.append(c)

        group_summary = results_df.groupby("グループ")[display_cols].sum().reset_index()
        group_summary["完了率"] = (group_summary["完了"] / group_summary["受電数"] * 100).round(2)
        group_summary = group_summary.sort_values("受電数", ascending=False)

        col1, col2 = st.columns(2)
        with col1:
            fig = px.bar(group_summary, x="グループ", y="受電数", title="グループ別 受電数",
                         color="受電数", color_continuous_scale="Blues", text="受電数")
            fig.update_traces(textposition="outside")
            st.plotly_chart(fig, use_container_width=True)
        with col2:
            fig = px.bar(group_summary, x="グループ", y="完了率", title="グループ別 完了率（%）",
                         color="完了率", color_continuous_scale="Greens", text="完了率")
            fig.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
            fig.update_yaxes(range=[0, 100])
            st.plotly_chart(fig, use_container_width=True)

        # グループ別テーブル
        gs_display = group_summary.copy()
        gs_display["完了率"] = gs_display["完了率"].apply(lambda x: f"{x:.1f}%")
        st.dataframe(gs_display, use_container_width=True, hide_index=True)

        # --- 担当者別 ---
        st.subheader("担当者別実績")
        group_filter = st.selectbox("グループで絞り込み", ["全て"] + list(groups.keys()) + ["未割当"])
        if group_filter != "全て":
            filtered = results_df[results_df["グループ"] == group_filter]
        else:
            filtered = results_df

        col1, col2 = st.columns(2)
        with col1:
            fig = px.bar(
                filtered.sort_values("受電数", ascending=True),
                x="受電数", y="担当者", orientation="h",
                title="担当者別 受電数", color="グループ",
                text="受電数",
            )
            fig.update_traces(textposition="outside")
            fig.update_layout(height=max(400, len(filtered) * 30))
            st.plotly_chart(fig, use_container_width=True)
        with col2:
            fig = px.bar(
                filtered.sort_values("完了率", ascending=True),
                x="完了率", y="担当者", orientation="h",
                title="担当者別 完了率（%）", color="グループ",
                text="完了率",
            )
            fig.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
            fig.update_layout(height=max(400, len(filtered) * 30))
            st.plotly_chart(fig, use_container_width=True)

        # テーブル
        st.subheader("全担当者データ")
        table_cols = ["グループ", "担当者", "受電数", "完了", "完了率"] + [
            c for c in optional_cols if c in filtered.columns
        ]
        table_df = filtered[table_cols].sort_values("受電数", ascending=False).copy()
        table_df["完了率"] = table_df["完了率"].apply(lambda x: f"{x:.1f}%")
        st.dataframe(table_df, use_container_width=True, hide_index=True)


# ----- タブ3: 稼働実績 -----
with tab_shift:
    st.subheader("CS入電 稼働実績")

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
                {"日": f"{d}日", "出勤者数": c} for d, c in sorted(daily_counts.items())
            ])
            fig = px.bar(dc_df, x="日", y="出勤者数", title="日別 出勤者数",
                         color="出勤者数", color_continuous_scale="Purples", text="出勤者数")
            fig.update_traces(textposition="outside")
            st.plotly_chart(fig, use_container_width=True)

        col1, col2 = st.columns(2)
        with col1:
            s = shift_df.sort_values("稼働日数", ascending=True)
            fig = px.bar(s, x="稼働日数", y="担当者", title="担当者別 稼働日数",
                         orientation="h", color="稼働日数", color_continuous_scale="Purples")
            fig.update_layout(height=max(400, len(s) * 25))
            st.plotly_chart(fig, use_container_width=True)
        with col2:
            s = shift_df.sort_values("実績時間", ascending=True)
            fig = go.Figure()
            fig.add_trace(go.Bar(y=s["担当者"], x=s["予定時間"], name="予定時間",
                                 orientation="h", marker_color="#BBDEFB"))
            fig.add_trace(go.Bar(y=s["担当者"], x=s["実績時間"], name="実績時間",
                                 orientation="h", marker_color="#1565C0"))
            fig.update_layout(title="予定時間 vs 実績時間", barmode="overlay",
                              height=max(400, len(s) * 25))
            st.plotly_chart(fig, use_container_width=True)

        # 日別 受電率 × 稼働時間
        if daily_counts and not daily_df.empty:
            st.subheader("日別 受電率 × 稼働時間")
            dr = daily_df[["日付", "曜日", "入電数", "受電対応数", "受電率"]].copy()
            dr["日"] = dr["日付"].dt.day

            # 日別の合計稼働時間を計算
            daily_hours = {}
            for _, row in shift_df.iterrows():
                for day, hours in row["日別稼働"].items():
                    daily_hours[day] = daily_hours.get(day, 0) + hours
            dh = pd.DataFrame([{"日": d, "稼働時間合計": round(h, 1)} for d, h in daily_hours.items()])

            dc = pd.DataFrame([{"日": d, "出勤者数": c} for d, c in daily_counts.items()])
            merged = pd.merge(dr, dc, on="日", how="inner")
            merged = pd.merge(merged, dh, on="日", how="left").fillna(0)
            merged["日付_str"] = merged["日付"].dt.strftime("%m/%d") + "(" + merged["曜日"] + ")"

            fig = go.Figure()
            # 稼働時間（棒グラフ）
            fig.add_trace(go.Bar(
                x=merged["日付_str"], y=merged["稼働時間合計"],
                name="稼働時間合計(h)", marker_color="#CE93D8", opacity=0.5,
                yaxis="y2",
            ))
            # 受電率（折れ線）
            fig.add_trace(go.Scatter(
                x=merged["日付_str"], y=merged["受電率"],
                name="受電率(%)", mode="lines+markers",
                line=dict(color="#1565C0", width=3),
                yaxis="y1",
            ))
            # 入電数（折れ線）
            fig.add_trace(go.Scatter(
                x=merged["日付_str"], y=merged["入電数"],
                name="入電数", mode="lines+markers",
                line=dict(color="#FF7043", width=2, dash="dot"),
                yaxis="y3",
            ))
            fig.update_layout(
                title="日別 受電率 × 稼働時間 × 入電数",
                xaxis=dict(tickangle=45),
                yaxis=dict(title="受電率（%）", side="left", range=[0, 100]),
                yaxis2=dict(title="稼働時間(h)", overlaying="y", side="right"),
                yaxis3=dict(overlaying="y", side="right", showticklabels=False),
                legend=dict(x=0, y=1.15, orientation="h"),
                height=500,
            )
            st.plotly_chart(fig, use_container_width=True)

            # テーブルも表示
            table = merged[["日付_str", "入電数", "受電対応数", "受電率", "出勤者数", "稼働時間合計"]].copy()
            table.columns = ["日付", "入電数", "受電対応数", "受電率(%)", "出勤者数", "稼働時間合計(h)"]
            table["受電率(%)"] = table["受電率(%)"].apply(lambda x: f"{x:.1f}%")
            st.dataframe(table, use_container_width=True, hide_index=True)

        st.subheader("稼働実績一覧")
        ds = shift_df[["担当者", "稼働日数", "予定時間", "実績時間"]].sort_values("実績時間", ascending=False)
        st.dataframe(ds, use_container_width=True, hide_index=True)


# ----- タブ4: 改善分析 -----
with tab_improve:
    st.subheader("🎯 受電率 改善分析")
    st.caption("入電数・受電率・稼働状況を突き合わせて、どこにスタッフを配置すべきかを分析します。")

    if hourly_df.empty or shift_df.empty or daily_df.empty:
        st.warning("データが不足しています（受電率・稼働実績の両方が必要です）")
    else:
        # --- 1. 時間帯別: 平均入電数 vs 平均受電率 ---
        st.subheader("時間帯別 入電負荷 × 受電率")
        st.caption("受電率が低く入電数が多い時間帯 = スタッフ増員の優先度が高い")

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

            fig = go.Figure()
            # 平均入電数（棒グラフ）
            fig.add_trace(go.Bar(
                x=h_stats["時間帯"], y=h_stats["平均入電数"],
                name="平均入電数", marker_color="#BBDEFB", opacity=0.7,
                yaxis="y2",
            ))
            # 取りこぼし数（棒グラフ）
            fig.add_trace(go.Bar(
                x=h_stats["時間帯"], y=h_stats["取りこぼし"],
                name="取りこぼし（月合計）", marker_color="#EF9A9A", opacity=0.7,
                yaxis="y3",
            ))
            # 受電率（折れ線）
            fig.add_trace(go.Scatter(
                x=h_stats["時間帯"], y=h_stats["平均受電率"],
                name="平均受電率(%)", mode="lines+markers+text",
                line=dict(color="#1565C0", width=3),
                text=[f"{v:.1f}%" for v in h_stats["平均受電率"]],
                textposition="top center",
                yaxis="y1",
            ))
            fig.update_layout(
                title="時間帯別 平均入電数 × 受電率 × 取りこぼし",
                yaxis=dict(title="受電率（%）", side="left", range=[0, 100]),
                yaxis2=dict(title="平均入電数", overlaying="y", side="right"),
                yaxis3=dict(overlaying="y", side="right", showticklabels=False),
                legend=dict(x=0, y=1.15, orientation="h"),
                barmode="group",
                height=500,
            )
            st.plotly_chart(fig, use_container_width=True)

            # 優先度テーブル
            priority = h_stats[["時間帯", "平均入電数", "平均受電率", "取りこぼし"]].copy()
            priority["優先度"] = priority.apply(
                lambda r: "🔴 高" if r["平均受電率"] < 80 and r["平均入電数"] > priority["平均入電数"].median()
                else ("🟡 中" if r["平均受電率"] < 90 else "🟢 低"), axis=1
            )
            priority["平均受電率"] = priority["平均受電率"].apply(lambda x: f"{x:.1f}%")
            st.dataframe(priority, use_container_width=True, hide_index=True)

        # --- 2. 日別: 出勤者数 × 受電率 の相関 ---
        st.subheader("日別 出勤者数 × 受電率 の関係")
        st.caption("出勤者数と受電率の関係から、必要な最低人数の目安がわかります")

        daily_counts = shift_df.attrs.get("daily_staff_count", {})
        if daily_counts:
            dr = daily_df[["日付", "曜日", "入電数", "受電対応数", "受電率"]].copy()
            dr["日"] = dr["日付"].dt.day
            dc = pd.DataFrame([{"日": d, "出勤者数": c} for d, c in daily_counts.items()])
            merged = pd.merge(dr, dc, on="日", how="inner")

            # 散布図
            fig = px.scatter(
                merged, x="出勤者数", y="受電率",
                size="入電数", color="曜日",
                hover_data=["日付", "入電数", "受電対応数"],
                title="出勤者数 vs 受電率（バブルサイズ = 入電数）",
                size_max=30,
            )
            fig.update_yaxes(range=[0, 100])
            # 目標ライン
            fig.add_hline(y=90, line_dash="dash", line_color="green",
                          annotation_text="目標 90%", annotation_position="top left")
            fig.update_layout(height=500)
            st.plotly_chart(fig, use_container_width=True)

            # 相関の簡易分析
            corr = merged["出勤者数"].corr(merged["受電率"])
            if corr > 0.3:
                st.info(f"📊 出勤者数と受電率の相関係数: **{corr:.2f}**（正の相関あり → スタッフ増で受電率向上が期待できます）")
            elif corr > 0:
                st.info(f"📊 出勤者数と受電率の相関係数: **{corr:.2f}**（弱い正の相関 → スタッフ数以外の要因も大きい可能性があります）")
            else:
                st.info(f"📊 出勤者数と受電率の相関係数: **{corr:.2f}**（相関なし → 入電数の変動や時間帯配置の問題が大きい可能性があります）")

            # 受電率90%以上の日 vs 未満の日
            good = merged[merged["受電率"] >= 90]
            bad = merged[merged["受電率"] < 90]
            if not good.empty and not bad.empty:
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("受電率 90%以上の日の平均出勤者数", f"{good['出勤者数'].mean():.1f}名")
                    st.metric("受電率 90%以上の日の平均入電数", f"{good['入電数'].mean():.0f}件")
                with col2:
                    st.metric("受電率 90%未満の日の平均出勤者数", f"{bad['出勤者数'].mean():.1f}名")
                    st.metric("受電率 90%未満の日の平均入電数", f"{bad['入電数'].mean():.0f}件")

            # テーブル
            st.subheader("日別データ一覧")
            table = merged[["日付", "曜日", "入電数", "受電対応数", "受電率", "出勤者数"]].copy()
            table["日付"] = table["日付"].dt.strftime("%m/%d")
            table["受電率"] = table["受電率"].apply(lambda x: f"{x:.1f}%")
            table = table.sort_values("日付")
            st.dataframe(table, use_container_width=True, hide_index=True)


# ----- タブ5: グループ設定 -----
with tab_groups:
    st.subheader("⚙️ グループ設定")
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
            st.success("グループ設定を保存しました！")
            st.rerun()
    with col2:
        if st.button("🔄 デフォルトに戻す"):
            save_groups(DEFAULT_GROUPS)
            st.session_state.groups = DEFAULT_GROUPS
            st.success("デフォルト設定に戻しました！")
            st.rerun()

    # 現在の設定表示
    st.markdown("---")
    st.subheader("現在の設定")
    for g, members in groups.items():
        st.markdown(f"**{g}**: {', '.join(members)}")
