# CS入電 チーム分析ダッシュボード

## プロジェクト概要
CS（カスタマーサポート）入電チームの分析ダッシュボード。Salesforceからデータを取得し、Streamlitで可視化する。

## 起動方法
```bash
streamlit run app.py
```

## 技術スタック
- **フロントエンド**: Streamlit + Plotly
- **データ取得**: simple-salesforce (Salesforce API)
- **データ処理**: pandas
- **AIエージェント**: anthropic SDK (Claude API) - `agent.py`
- **Python依存**: `requirements.txt` (`pip install -r requirements.txt`)

## ファイル構成
- `app.py` - メインのStreamlitダッシュボード（4タブ: 受電率、コール処理実績、稼働実績、グループ設定）
- `salesforce_client.py` - Salesforce APIクライアント（データ取得ロジック全般）
- `agent.py` - Claude Agent SDKを使った自然言語分析エージェント（開発中）
- `groups.json` - グループ設定の保存ファイル（gitignore対象）
- `.env` - 環境変数（gitignore対象）

## Salesforceデータソース
- `ZVC__Zoom_Call_Log__c` - Zoom通話ログ（受電率計算用、Call_IDで一意着信を特定）
- `Task` (Status='clok', Field3_del__c='受電') - コール処理実績
- `CustomObject11__c` - 稼働実績（シフト）
- `CustomObject10__c` - 人事情報（CS部署スタッフ）

## ビジネスルール
- 営業時間: JST 10:00-19:00
- UTCフィルター: `HOUR_IN_DAY(CreatedDate) >= 1 AND HOUR_IN_DAY(CreatedDate) < 10`
- キャッシュ: 30分 (`st.cache_data(ttl=1800)`)
- グループ: O既存、Z中堅、Z新人、O新人、バイトル（UI上で編集・並び替え可能）

## 環境変数
- `SF_USERNAME`, `SF_PASSWORD`, `SF_SECURITY_TOKEN`, `SF_DOMAIN` - Salesforce認証
- `ANTHROPIC_API_KEY` - Claude API（agent.py用）

## 開発上の注意
- 日本語でコミットメッセージを書く
- `.env`、`groups.json`、`*.json`（requirements.txt除く）はgitignore対象
- Streamlit Cloud デプロイ時は `st.secrets` から認証情報を取得
