# KatagrMa 保育経営診断ツール

AI分析による保育施設経営診断。Beyond SaaS 戦略の Step 1「まず懐に入る」を実現する入口ツール。

Claude Opus 4.7 による所見生成、SQLite 永続化、Zoho CRM 自動連携、メール配信、Basic Auth 管理画面を実装。

## ローカル起動

```bash
npm install
cp .env.example .env     # ANTHROPIC_API_KEY を設定
npm start                # http://localhost:3000/shindan.html
```

## 環境変数

| 変数 | 必須 | 用途 |
|------|------|------|
| `ANTHROPIC_API_KEY` | ✅ | Claude API 認証 |
| `ADMIN_PASSWORD` | — | `/admin` の Basic Auth（未設定時は `/admin` が 404） |
| `ZOHO_WEBHOOK_URL` | — | Zoho CRM リード自動登録 |
| `RESEND_API_KEY` | — | サンキュー/営業通知メール送信 |
| `FROM_EMAIL` | — | メール送信元（Resend で認証済みドメイン） |
| `SALES_EMAIL` | — | 相談申込時の営業チーム通知先 |
| `PORT` | — | サーバーポート（デフォルト 3000） |

## エンドポイント

| Method | Path | 用途 |
|--------|------|------|
| GET | `/shindan.html` | 診断ページ |
| POST | `/api/analyze` | AI 分析（Claude Opus 4.7） |
| POST | `/api/save` | 診断結果を SQLite に保存 + Zoho/メール配信 |
| POST | `/api/consultation` | 個別相談申込 |
| GET | `/admin` | 管理画面（Basic Auth） |
| GET | `/admin/:id` | 診断詳細 |
| GET | `/health` | ヘルスチェック |

## Render.com デプロイ

1. このリポジトリを GitHub に push
2. Render で **New Web Service** → GitHub リポジトリを選択
3. Build Command: `npm install`、Start Command: `npm start`
4. Environment タブで `ANTHROPIC_API_KEY` と `ADMIN_PASSWORD` を設定
5. Deploy

`render.yaml` を配置済みのため、**Blueprint** 経由で一括セットアップも可能。

> **Note:** Free プランは永続ディスクがないため、再起動で SQLite データは消えます。本番運用では有料プラン（Starter 以上）もしくは外部 DB（Supabase / Turso）への移行を推奨。

## マネジメント共有資料

- [KatagrMa_Shindan_Report.pptx](KatagrMa_Shindan_Report.pptx) — プロトタイプ報告デッキ（8 枚）
- [KatagrMa_Shindan_Report.pdf](KatagrMa_Shindan_Report.pdf) — 同 PDF 版

## Phase ロードマップ

- **Phase 1 (完了)** — 診断 UI、スコアリング、AI 分析、結果レポート
- **Phase 2 (完了)** — SQLite 永続化、Zoho Webhook、メール配信、管理画面
- **Phase 3 (着手予定)** — 質問 30 問の最終化、535 施設ベンチマーク算出、本番デプロイ、ウェビナー連携
