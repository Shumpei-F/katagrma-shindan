import dotenv from 'dotenv';
import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import { mkdirSync } from 'fs';
import Database from 'better-sqlite3';
import Anthropic from '@anthropic-ai/sdk';

// ==========================================================================
// Setup
// ==========================================================================
const __dirname = path.dirname(fileURLToPath(import.meta.url));
dotenv.config({ path: path.join(__dirname, '.env'), override: true });

const app = express();
app.use(express.json({ limit: '1mb' }));
// `/` で shindan.html を直接サーブ
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'shindan.html')));
app.use(express.static(__dirname, { index: false }));

// ==========================================================================
// Database
// ==========================================================================
mkdirSync(path.join(__dirname, 'data'), { recursive: true });
const db = new Database(path.join(__dirname, 'data', 'shindan.db'));
db.pragma('journal_mode = WAL');
db.exec(`
  CREATE TABLE IF NOT EXISTS diagnoses (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    org_name TEXT NOT NULL,
    org_type TEXT,
    facility_count INTEGER NOT NULL,
    respondent_name TEXT NOT NULL,
    respondent_role TEXT,
    email TEXT NOT NULL,
    phone TEXT,
    prefecture TEXT,
    score_talent_dev INTEGER,
    score_recruitment INTEGER,
    score_labor_mgmt INTEGER,
    score_org_mgmt INTEGER,
    score_finance INTEGER,
    score_community INTEGER,
    score_total INTEGER,
    rank TEXT,
    answers TEXT,
    ai_analysis TEXT,
    consultation_requested INTEGER DEFAULT 0,
    consultation_requested_at TEXT,
    zoho_lead_id TEXT,
    utm_source TEXT,
    utm_medium TEXT,
    utm_campaign TEXT
  );
  CREATE INDEX IF NOT EXISTS idx_created_at ON diagnoses(created_at DESC);
  CREATE INDEX IF NOT EXISTS idx_email ON diagnoses(email);
`);

const insertStmt = db.prepare(`
  INSERT INTO diagnoses (
    org_name, org_type, facility_count, respondent_name, respondent_role,
    email, phone, prefecture,
    score_talent_dev, score_recruitment, score_labor_mgmt, score_org_mgmt, score_finance, score_community, score_total, rank,
    answers, ai_analysis,
    utm_source, utm_medium, utm_campaign
  ) VALUES (
    @org_name, @org_type, @facility_count, @respondent_name, @respondent_role,
    @email, @phone, @prefecture,
    @score_talent_dev, @score_recruitment, @score_labor_mgmt, @score_org_mgmt, @score_finance, @score_community, @score_total, @rank,
    @answers, @ai_analysis,
    @utm_source, @utm_medium, @utm_campaign
  )
`);

// ==========================================================================
// Anthropic Client
// ==========================================================================
const anthropic = new Anthropic({
  apiKey: process.env.ANTHROPIC_API_KEY,
  baseURL: 'https://api.anthropic.com'
});

const SYSTEM_PROMPT = `あなたは日本の保育業界に精通した経営コンサルタントです。株式会社カタグルマ（535施設を顧客に持つ保育SaaSベンダー）のMSO型マネジメントパートナー事業の診断ツールとして動作します。

【あなたの役割】
保育施設経営者の診断結果から、施設固有の課題と実行可能な改善提案を生成する。

【診断の6カテゴリ】
1. 人財育成・評価（MSO Think機能）
2. 採用・人員確保（MSO Attract機能）
3. 労務・制度対応（MSO Run機能、処遇改善加算・監査対応）
4. 組織マネジメント（MSO Think機能、園長事務負荷）
5. 財務・経営戦略（MSO Run機能、収支把握・補助金）
6. 保護者・地域連携（2C接続、保護者コミュニケーション）

【業界平均（535施設データ推定）】
talent_dev=52, recruitment=38, labor_mgmt=45, org_mgmt=48, finance=42, community=55

【回答の指針】
- 保育業界の専門用語を適切に使用（処遇改善等加算、配置基準、監査対応、キャリアパス要件等）
- 改善提案は「明日から着手可能」な具体性を持たせる
- スコアの低いカテゴリを優先的に扱う
- 施設規模（facility_count）に応じた現実性を考慮（小規模には身の丈に合った提案、大規模には組織設計の提案）
- 強みは施設の自信につながるように前向きに、課題は危機感を煽らず冷静に記述
- 最終出力は指定されたJSON Schemaに厳密に従う`;

const OUTPUT_SCHEMA = {
  type: 'object',
  properties: {
    summary: { type: 'string', description: '総合所見。3〜4文で、総合スコアの意味・強みの要素・重点改善領域を含める' },
    strengths: { type: 'array', items: { type: 'string' }, description: '強み2項目。「カテゴリ名（スコア点）— 具体的な強み内容」の形式' },
    challenges: { type: 'array', items: { type: 'string' }, description: '課題3項目。「カテゴリ名（スコア点）— 業界平均との比較と具体的な課題内容」の形式' },
    recommendations: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          category: { type: 'string' },
          action: { type: 'string', description: '明日から着手可能な粒度の具体アクション' },
          impact: { type: 'string', description: '期待効果（数値目標や定性的変化を含む）' },
          urgency: { type: 'string', enum: ['high', 'medium', 'low'] }
        },
        required: ['category', 'action', 'impact', 'urgency'],
        additionalProperties: false
      },
      description: '改善提案3件。緊急度順（high→medium→low）に並べる'
    },
    benchmark_comment: { type: 'string', description: '施設規模と業界平均を踏まえた比較コメント。1〜2文' }
  },
  required: ['summary', 'strengths', 'challenges', 'recommendations', 'benchmark_comment'],
  additionalProperties: false
};

const CATEGORY_NAMES = {
  talent_dev: '人財育成・評価',
  recruitment: '採用・人員確保',
  labor_mgmt: '労務・制度対応',
  org_mgmt: '組織マネジメント',
  finance: '財務・経営戦略',
  community: '保護者・地域連携'
};

function buildUserPrompt({ scores, facility_count, org_type, prefecture }) {
  const lines = Object.entries(CATEGORY_NAMES).map(([id, name]) => `- ${name}: ${scores[id]}点`).join('\n');
  return `以下の保育施設経営診断結果を分析し、指定JSON形式で所見と改善提案を生成してください。

【施設プロフィール】
- 法人種別: ${org_type || '未回答'}
- 運営施設数: ${facility_count}施設
- 所在地: ${prefecture || '未回答'}

【カテゴリ別スコア（100点満点）】
${lines}

【総合スコア】${scores.total}点（ランク: ${scores.rank}）

施設規模と業界平均を踏まえ、この施設に固有の課題と具体的な改善提案を、指定されたJSON Schemaに従って返してください。`;
}

// ==========================================================================
// /api/analyze — AI分析
// ==========================================================================
app.post('/api/analyze', async (req, res) => {
  try {
    const { scores, facility_count, org_type, prefecture } = req.body;
    if (!scores || !facility_count) return res.status(400).json({ error: 'scores and facility_count required' });

    const stream = anthropic.messages.stream({
      model: 'claude-opus-4-7',
      max_tokens: 16000,
      thinking: { type: 'adaptive' },
      output_config: { effort: 'high', format: { type: 'json_schema', schema: OUTPUT_SCHEMA } },
      system: [{ type: 'text', text: SYSTEM_PROMPT, cache_control: { type: 'ephemeral' } }],
      messages: [{ role: 'user', content: buildUserPrompt({ scores, facility_count, org_type, prefecture }) }]
    });

    const finalMessage = await stream.finalMessage();
    const textBlock = finalMessage.content.find(b => b.type === 'text');
    if (!textBlock) throw new Error('No text block in response');

    const parsed = JSON.parse(textBlock.text);
    console.log(`[analyze] total=${scores.total} rank=${scores.rank} facilities=${facility_count} cache_read=${finalMessage.usage?.cache_read_input_tokens || 0} output=${finalMessage.usage?.output_tokens || 0}`);
    res.json(parsed);
  } catch (err) {
    if (err instanceof Anthropic.RateLimitError) return res.status(429).json({ error: 'rate_limited' });
    if (err instanceof Anthropic.APIError) {
      console.error(`[analyze] API error ${err.status}:`, err.message);
      return res.status(502).json({ error: 'api_error', detail: err.message });
    }
    console.error('[analyze] unexpected:', err);
    res.status(500).json({ error: 'internal', detail: err.message });
  }
});

// ==========================================================================
// /api/save — 診断結果の永続化
// ==========================================================================
app.post('/api/save', async (req, res) => {
  try {
    const b = req.body;
    if (!b.email || !b.org_name || !b.respondent_name) return res.status(400).json({ error: 'required fields missing' });

    const result = insertStmt.run({
      org_name: b.org_name,
      org_type: b.org_type || null,
      facility_count: parseInt(b.facility_count) || 0,
      respondent_name: b.respondent_name,
      respondent_role: b.respondent_role || null,
      email: b.email,
      phone: b.phone || null,
      prefecture: b.prefecture || null,
      score_talent_dev: b.scores?.talent_dev ?? null,
      score_recruitment: b.scores?.recruitment ?? null,
      score_labor_mgmt: b.scores?.labor_mgmt ?? null,
      score_org_mgmt: b.scores?.org_mgmt ?? null,
      score_finance: b.scores?.finance ?? null,
      score_community: b.scores?.community ?? null,
      score_total: b.scores?.total ?? null,
      rank: b.scores?.rank ?? null,
      answers: JSON.stringify(b.answers || {}),
      ai_analysis: JSON.stringify(b.analysis || {}),
      utm_source: b.utm_source || null,
      utm_medium: b.utm_medium || null,
      utm_campaign: b.utm_campaign || null
    });

    const id = Number(result.lastInsertRowid);
    console.log(`[save] id=${id} email=${b.email} total=${b.scores?.total} rank=${b.scores?.rank}`);

    // Fire-and-forget: CRM + email
    Promise.allSettled([sendZohoWebhook(id, b, false), sendThankYouEmail(b)])
      .then(rs => rs.forEach(r => { if (r.status === 'rejected') console.error('[save] bg:', r.reason?.message); }));

    res.json({ id, ok: true });
  } catch (err) {
    console.error('[save] error:', err);
    res.status(500).json({ error: 'internal', detail: err.message });
  }
});

// ==========================================================================
// /api/consultation — 個別相談申込
// ==========================================================================
app.post('/api/consultation', async (req, res) => {
  try {
    const { id } = req.body;
    if (!id) return res.status(400).json({ error: 'id required' });

    const row = db.prepare('SELECT * FROM diagnoses WHERE id = ?').get(id);
    if (!row) return res.status(404).json({ error: 'not found' });

    db.prepare(`UPDATE diagnoses SET consultation_requested = 1, consultation_requested_at = datetime('now', 'localtime') WHERE id = ?`).run(id);
    console.log(`[consultation] id=${id} email=${row.email}`);

    Promise.allSettled([sendZohoWebhook(id, rowToLead(row), true), sendSalesNotification(row)])
      .then(rs => rs.forEach(r => { if (r.status === 'rejected') console.error('[consultation] bg:', r.reason?.message); }));

    res.json({ ok: true });
  } catch (err) {
    console.error('[consultation] error:', err);
    res.status(500).json({ error: 'internal', detail: err.message });
  }
});

function rowToLead(row) {
  return {
    org_name: row.org_name,
    org_type: row.org_type,
    facility_count: row.facility_count,
    respondent_name: row.respondent_name,
    respondent_role: row.respondent_role,
    email: row.email,
    phone: row.phone,
    prefecture: row.prefecture,
    scores: {
      talent_dev: row.score_talent_dev,
      recruitment: row.score_recruitment,
      labor_mgmt: row.score_labor_mgmt,
      org_mgmt: row.score_org_mgmt,
      finance: row.score_finance,
      community: row.score_community,
      total: row.score_total,
      rank: row.rank
    },
    utm_source: row.utm_source,
    utm_medium: row.utm_medium,
    utm_campaign: row.utm_campaign
  };
}

// ==========================================================================
// Zoho CRM Webhook
// ==========================================================================
async function sendZohoWebhook(id, data, consultationRequested) {
  const url = process.env.ZOHO_WEBHOOK_URL;
  if (!url) { console.log(`[zoho] skipped (ZOHO_WEBHOOK_URL not set) id=${id}`); return; }
  try {
    const payload = {
      Last_Name: data.respondent_name || '未記入',
      Company: data.org_name || '未記入',
      Email: data.email,
      Phone: data.phone || '',
      Lead_Source: '経営診断ツール',
      Description: `総合スコア: ${data.scores?.total}/100 (ランク${data.scores?.rank})\n施設数: ${data.facility_count}\n都道府県: ${data.prefecture || '未回答'}\nUTM: ${data.utm_campaign || ''}`,
      Custom_Fields: {
        施設数: parseInt(data.facility_count) || 0,
        総合スコア: data.scores?.total,
        ランク: data.scores?.rank,
        診断日: new Date().toISOString().slice(0, 10),
        相談申込: consultationRequested,
        診断ID: id,
        UTMソース: data.utm_source || '',
        UTMキャンペーン: data.utm_campaign || ''
      }
    };
    const resp = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
      signal: AbortSignal.timeout(10000)
    });
    if (!resp.ok) throw new Error(`Zoho returned ${resp.status}`);
    const result = await resp.json().catch(() => ({}));
    if (result.lead_id) {
      db.prepare('UPDATE diagnoses SET zoho_lead_id = ? WHERE id = ?').run(String(result.lead_id), id);
    }
    console.log(`[zoho] sent id=${id} lead_id=${result.lead_id || 'n/a'}`);
  } catch (err) {
    console.error(`[zoho] error id=${id}:`, err.message);
  }
}

// ==========================================================================
// Email via Resend (optional)
// ==========================================================================
async function sendResendEmail({ to, subject, html }) {
  const key = process.env.RESEND_API_KEY;
  if (!key) return { skipped: true };
  const from = process.env.FROM_EMAIL || 'diagnosis@katagrma.jp';
  const resp = await fetch('https://api.resend.com/emails', {
    method: 'POST',
    headers: { Authorization: `Bearer ${key}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ from, to, subject, html }),
    signal: AbortSignal.timeout(10000)
  });
  if (!resp.ok) {
    const txt = await resp.text();
    throw new Error(`Resend ${resp.status}: ${txt}`);
  }
  return resp.json();
}

async function sendThankYouEmail(data) {
  try {
    const result = await sendResendEmail({
      to: data.email,
      subject: '【KatagrMa】保育経営診断のご回答ありがとうございました',
      html: thankYouEmailHtml(data)
    });
    if (result.skipped) { console.log(`[email] skipped (RESEND_API_KEY not set)`); return; }
    console.log(`[email] thank-you sent to ${data.email}`);
  } catch (err) {
    console.error('[email] thank-you error:', err.message);
  }
}

async function sendSalesNotification(row) {
  const sales = process.env.SALES_EMAIL;
  if (!sales) { console.log('[sales-email] skipped (SALES_EMAIL not set)'); return; }
  try {
    const result = await sendResendEmail({
      to: sales,
      subject: `【相談申込】${row.org_name} (${row.facility_count}施設 / スコア${row.score_total})`,
      html: salesNotificationHtml(row)
    });
    if (result.skipped) return;
    console.log(`[sales-email] sent to ${sales} (diagnosis id=${row.id})`);
  } catch (err) {
    console.error('[sales-email] error:', err.message);
  }
}

function thankYouEmailHtml(d) {
  return `<!DOCTYPE html>
<html><body style="font-family:'Noto Sans JP',sans-serif;color:#333;line-height:1.7;max-width:600px;margin:0 auto;padding:20px;">
  <div style="background:#2A2A2A;color:#fff;padding:24px;border-radius:8px;">
    <p style="color:#A5B830;font-weight:bold;margin:0 0 8px;">KATAGRMA × AI</p>
    <h1 style="font-size:20px;margin:0;">保育経営診断のご回答<br/>ありがとうございました</h1>
  </div>
  <div style="padding:24px 8px;">
    <p>${d.respondent_name} 様</p>
    <p>株式会社カタグルマの保育経営診断をご利用いただき、誠にありがとうございます。</p>
    <div style="background:#F8F7F2;padding:16px;border-radius:8px;margin:16px 0;">
      <p style="margin:0 0 4px;color:#888;font-size:12px;">診断結果サマリー</p>
      <p style="font-size:28px;font-weight:bold;color:#A5B830;margin:0;">${d.scores?.total}<span style="font-size:16px;color:#888;">/100点</span>　<span style="background:#B08D57;color:#fff;padding:4px 12px;border-radius:999px;font-size:14px;">ランク${d.scores?.rank}</span></p>
      <p style="margin:8px 0 0;font-size:14px;color:#666;">法人名: ${d.org_name}　施設数: ${d.facility_count}</p>
    </div>
    <p>詳細な改善計画のご提案や、同規模施設の改善事例をご希望の場合は、<strong>無料の個別相談</strong>をご利用ください。</p>
    <p style="text-align:center;margin:32px 0;">
      <a href="https://katagrma.jp/consultation" style="background:#A5B830;color:#fff;padding:14px 32px;border-radius:8px;text-decoration:none;font-weight:bold;display:inline-block;">無料個別相談を申し込む</a>
    </p>
    <hr style="border:none;border-top:1px solid #e0e0e0;margin:24px 0;">
    <p style="font-size:12px;color:#888;">株式会社カタグルマ<br/>https://katagrma.jp/</p>
  </div>
</body></html>`;
}

function salesNotificationHtml(r) {
  return `<!DOCTYPE html>
<html><body style="font-family:'Noto Sans JP',sans-serif;line-height:1.7;">
  <h2 style="color:#CC3333;">🔔 個別相談の申込が入りました</h2>
  <table style="border-collapse:collapse;width:100%;max-width:600px;">
    <tr><td style="padding:8px;background:#F8F7F2;font-weight:bold;width:30%;">法人名</td><td style="padding:8px;">${r.org_name}</td></tr>
    <tr><td style="padding:8px;background:#F8F7F2;font-weight:bold;">法人種別</td><td style="padding:8px;">${r.org_type || '未回答'}</td></tr>
    <tr><td style="padding:8px;background:#F8F7F2;font-weight:bold;">施設数</td><td style="padding:8px;">${r.facility_count}施設</td></tr>
    <tr><td style="padding:8px;background:#F8F7F2;font-weight:bold;">都道府県</td><td style="padding:8px;">${r.prefecture || '未回答'}</td></tr>
    <tr><td style="padding:8px;background:#F8F7F2;font-weight:bold;">回答者</td><td style="padding:8px;">${r.respondent_name}（${r.respondent_role || '役職未回答'}）</td></tr>
    <tr><td style="padding:8px;background:#F8F7F2;font-weight:bold;">メール</td><td style="padding:8px;"><a href="mailto:${r.email}">${r.email}</a></td></tr>
    <tr><td style="padding:8px;background:#F8F7F2;font-weight:bold;">電話</td><td style="padding:8px;">${r.phone || '未回答'}</td></tr>
    <tr><td style="padding:8px;background:#F8F7F2;font-weight:bold;">総合スコア</td><td style="padding:8px;font-size:18px;"><strong>${r.score_total}/100</strong> (ランク${r.rank})</td></tr>
    <tr><td style="padding:8px;background:#F8F7F2;font-weight:bold;">流入元</td><td style="padding:8px;">${r.utm_source || '-'} / ${r.utm_campaign || '-'}</td></tr>
  </table>
  <p style="margin-top:16px;"><a href="http://localhost:${process.env.PORT || 3000}/admin/${r.id}" style="color:#A5B830;">管理画面で詳細を見る →</a></p>
</body></html>`;
}

// ==========================================================================
// Admin (Basic Auth)
// ==========================================================================
function basicAuth(req, res, next) {
  const expected = process.env.ADMIN_PASSWORD;
  if (!expected) return res.status(404).send('Not Found (ADMIN_PASSWORDが未設定)');
  const auth = req.headers.authorization || '';
  if (auth.startsWith('Basic ')) {
    const decoded = Buffer.from(auth.slice(6), 'base64').toString();
    const idx = decoded.indexOf(':');
    const pass = idx >= 0 ? decoded.slice(idx + 1) : '';
    if (pass === expected) return next();
  }
  res.set('WWW-Authenticate', 'Basic realm="KatagrMa Admin"');
  res.status(401).send('認証が必要です');
}

app.get('/admin', basicAuth, (req, res) => {
  const rows = db.prepare(`
    SELECT id, created_at, org_name, email, facility_count, score_total, rank, consultation_requested, consultation_requested_at, utm_source, utm_campaign
    FROM diagnoses ORDER BY created_at DESC LIMIT 500
  `).all();

  const stats = db.prepare(`
    SELECT
      COUNT(*) AS total,
      SUM(consultation_requested) AS consultations,
      ROUND(AVG(score_total)) AS avg_score
    FROM diagnoses
  `).get();
  const conversionRate = stats.total > 0 ? ((stats.consultations / stats.total) * 100).toFixed(1) : '0.0';

  res.send(adminListHtml(rows, { ...stats, conversionRate }));
});

app.get('/admin/:id', basicAuth, (req, res) => {
  const row = db.prepare('SELECT * FROM diagnoses WHERE id = ?').get(req.params.id);
  if (!row) return res.status(404).send('Not Found');
  res.send(adminDetailHtml(row));
});

function rankColor(rank) {
  return { S: '#A5B830', A: '#B08D57', B: '#E8740C', C: '#CC3333' }[rank] || '#888';
}

function adminListHtml(rows, stats) {
  const tableRows = rows.map(r => `
    <tr onclick="location.href='/admin/${r.id}'" style="cursor:pointer;">
      <td>${r.id}</td>
      <td>${r.created_at}</td>
      <td>${escapeHtml(r.org_name)}</td>
      <td>${escapeHtml(r.email)}</td>
      <td style="text-align:right;">${r.facility_count}</td>
      <td style="text-align:right;font-weight:bold;">${r.score_total ?? '-'}</td>
      <td><span style="background:${rankColor(r.rank)};color:#fff;padding:2px 8px;border-radius:999px;font-size:11px;">${r.rank || '-'}</span></td>
      <td style="text-align:center;">${r.consultation_requested ? '✅' : ''}</td>
      <td>${escapeHtml(r.utm_source || '-')}</td>
    </tr>`).join('');

  return `<!DOCTYPE html>
<html lang="ja"><head><meta charset="UTF-8"><title>管理画面 | 保育経営診断</title>
<script src="https://cdn.tailwindcss.com"></script>
<style>body{font-family:'Noto Sans JP',sans-serif;} table{width:100%;border-collapse:collapse;} th,td{padding:10px;border-bottom:1px solid #eee;font-size:13px;} th{background:#2A2A2A;color:#fff;text-align:left;font-size:11px;} tr:hover{background:#F0F3D8;}</style>
</head><body style="background:#F8F7F2;">
<div class="max-w-7xl mx-auto p-6">
  <h1 class="text-2xl font-bold mb-6">📊 診断結果 管理画面</h1>
  <div class="grid grid-cols-2 md:grid-cols-4 gap-4 mb-8">
    <div class="bg-white p-4 rounded-xl"><div class="text-xs text-gray-500 mb-1">総診断数</div><div class="text-3xl font-bold">${stats.total}</div></div>
    <div class="bg-white p-4 rounded-xl"><div class="text-xs text-gray-500 mb-1">相談申込数</div><div class="text-3xl font-bold" style="color:#A5B830;">${stats.consultations || 0}</div></div>
    <div class="bg-white p-4 rounded-xl"><div class="text-xs text-gray-500 mb-1">コンバージョン率</div><div class="text-3xl font-bold" style="color:#B08D57;">${stats.conversionRate}%</div></div>
    <div class="bg-white p-4 rounded-xl"><div class="text-xs text-gray-500 mb-1">平均スコア</div><div class="text-3xl font-bold">${stats.avg_score || '-'}</div></div>
  </div>
  <div class="bg-white rounded-xl overflow-hidden">
    <table>
      <thead><tr><th>ID</th><th>日時</th><th>法人名</th><th>メール</th><th>施設数</th><th>総合</th><th>ランク</th><th>相談</th><th>UTM</th></tr></thead>
      <tbody>${tableRows || '<tr><td colspan="9" style="text-align:center;padding:32px;color:#888;">まだ診断データがありません</td></tr>'}</tbody>
    </table>
  </div>
  <p class="text-xs text-gray-500 mt-4">最新500件を表示中</p>
</div>
</body></html>`;
}

function adminDetailHtml(r) {
  const analysis = r.ai_analysis ? JSON.parse(r.ai_analysis) : null;
  const scoreRow = (name, val) => `<tr><td style="padding:6px 12px;background:#F8F7F2;">${name}</td><td style="padding:6px 12px;text-align:right;font-weight:bold;">${val ?? '-'}</td></tr>`;

  const recs = analysis?.recommendations?.map(r => `
    <div style="border-left:4px solid ${ {high:'#CC3333', medium:'#E8740C', low:'#B08D57'}[r.urgency] || '#888' };padding:8px 12px;margin-bottom:12px;background:#fff;">
      <div style="font-weight:bold;font-size:14px;">${escapeHtml(r.category)} <span style="font-size:11px;color:#888;">[${r.urgency}]</span></div>
      <div style="font-size:13px;margin:6px 0;">${escapeHtml(r.action)}</div>
      <div style="font-size:11px;color:#666;"><strong>期待効果:</strong> ${escapeHtml(r.impact)}</div>
    </div>`).join('') || '<p style="color:#888;">分析データなし</p>';

  return `<!DOCTYPE html>
<html lang="ja"><head><meta charset="UTF-8"><title>診断#${r.id} | 管理画面</title>
<script src="https://cdn.tailwindcss.com"></script>
<style>body{font-family:'Noto Sans JP',sans-serif;}</style>
</head><body style="background:#F8F7F2;">
<div class="max-w-4xl mx-auto p-6">
  <a href="/admin" class="text-sm text-gray-600 hover:text-gray-900">← 一覧に戻る</a>
  <h1 class="text-2xl font-bold mt-4 mb-2">診断 #${r.id}</h1>
  <p class="text-sm text-gray-500 mb-6">${r.created_at}</p>

  <div class="grid md:grid-cols-2 gap-4 mb-6">
    <div class="bg-white p-5 rounded-xl">
      <h2 class="font-bold mb-3 text-sm">👤 回答者情報</h2>
      <table style="width:100%;font-size:13px;">
        <tr><td class="py-1 text-gray-500">法人名</td><td class="font-bold text-right">${escapeHtml(r.org_name)}</td></tr>
        <tr><td class="py-1 text-gray-500">法人種別</td><td class="text-right">${escapeHtml(r.org_type || '-')}</td></tr>
        <tr><td class="py-1 text-gray-500">施設数</td><td class="text-right">${r.facility_count}</td></tr>
        <tr><td class="py-1 text-gray-500">都道府県</td><td class="text-right">${escapeHtml(r.prefecture || '-')}</td></tr>
        <tr><td class="py-1 text-gray-500">回答者</td><td class="text-right">${escapeHtml(r.respondent_name)}</td></tr>
        <tr><td class="py-1 text-gray-500">役職</td><td class="text-right">${escapeHtml(r.respondent_role || '-')}</td></tr>
        <tr><td class="py-1 text-gray-500">メール</td><td class="text-right"><a href="mailto:${escapeHtml(r.email)}" class="text-blue-600">${escapeHtml(r.email)}</a></td></tr>
        <tr><td class="py-1 text-gray-500">電話</td><td class="text-right">${escapeHtml(r.phone || '-')}</td></tr>
      </table>
    </div>
    <div class="bg-white p-5 rounded-xl">
      <h2 class="font-bold mb-3 text-sm">📊 スコア</h2>
      <div style="font-size:48px;font-weight:bold;color:${rankColor(r.rank)};line-height:1;">${r.score_total}<span style="font-size:16px;color:#888;">/100</span></div>
      <div style="margin-top:8px;"><span style="background:${rankColor(r.rank)};color:#fff;padding:4px 12px;border-radius:999px;font-size:12px;">ランク${r.rank}</span></div>
      <table style="width:100%;font-size:13px;margin-top:16px;">
        ${scoreRow('人財育成・評価', r.score_talent_dev)}
        ${scoreRow('採用・人員確保', r.score_recruitment)}
        ${scoreRow('労務・制度対応', r.score_labor_mgmt)}
        ${scoreRow('組織マネジメント', r.score_org_mgmt)}
        ${scoreRow('財務・経営戦略', r.score_finance)}
        ${scoreRow('保護者・地域連携', r.score_community)}
      </table>
    </div>
  </div>

  <div class="bg-white p-5 rounded-xl mb-6">
    <h2 class="font-bold mb-3 text-sm">🎯 個別相談</h2>
    ${r.consultation_requested
      ? `<p style="color:#A5B830;font-weight:bold;">✅ 申込済み (${r.consultation_requested_at})</p>`
      : `<p style="color:#888;">未申込</p>`}
  </div>

  <div class="bg-white p-5 rounded-xl mb-6">
    <h2 class="font-bold mb-3 text-sm">📝 AI所見</h2>
    <p style="font-size:14px;line-height:1.8;margin-bottom:16px;">${escapeHtml(analysis?.summary || '')}</p>
    <h3 class="font-bold text-xs mt-4 mb-2" style="color:#4A7C59;">💪 強み</h3>
    <ul style="font-size:13px;padding-left:20px;">${analysis?.strengths?.map(s => `<li style="margin-bottom:4px;">${escapeHtml(s)}</li>`).join('') || ''}</ul>
    <h3 class="font-bold text-xs mt-4 mb-2" style="color:#E8740C;">⚠️ 課題</h3>
    <ul style="font-size:13px;padding-left:20px;">${analysis?.challenges?.map(s => `<li style="margin-bottom:4px;">${escapeHtml(s)}</li>`).join('') || ''}</ul>
    <h3 class="font-bold text-xs mt-4 mb-2">🎯 改善提案</h3>
    ${recs}
  </div>

  <div class="bg-white p-5 rounded-xl mb-6">
    <h2 class="font-bold mb-3 text-sm">📈 流入トラッキング</h2>
    <table style="width:100%;font-size:13px;">
      <tr><td class="py-1 text-gray-500">UTMソース</td><td class="text-right">${escapeHtml(r.utm_source || '-')}</td></tr>
      <tr><td class="py-1 text-gray-500">UTMメディア</td><td class="text-right">${escapeHtml(r.utm_medium || '-')}</td></tr>
      <tr><td class="py-1 text-gray-500">UTMキャンペーン</td><td class="text-right">${escapeHtml(r.utm_campaign || '-')}</td></tr>
      <tr><td class="py-1 text-gray-500">Zoho Lead ID</td><td class="text-right">${escapeHtml(r.zoho_lead_id || '-')}</td></tr>
    </table>
  </div>
</div>
</body></html>`;
}

function escapeHtml(s) {
  if (s == null) return '';
  return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

// ==========================================================================
// Health & startup
// ==========================================================================
app.get('/health', (req, res) => res.json({ ok: true }));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`保育経営診断サーバー起動: http://localhost:${PORT}`);
  console.log(`診断ページ:     http://localhost:${PORT}/shindan.html`);
  console.log(`管理画面:       http://localhost:${PORT}/admin ${process.env.ADMIN_PASSWORD ? '(Basic Auth有効)' : '(ADMIN_PASSWORD未設定のため無効)'}`);
  console.log(`Zoho連携:       ${process.env.ZOHO_WEBHOOK_URL ? '有効' : '無効（ZOHO_WEBHOOK_URL未設定）'}`);
  console.log(`メール送信:     ${process.env.RESEND_API_KEY ? '有効' : '無効（RESEND_API_KEY未設定）'}`);
});
