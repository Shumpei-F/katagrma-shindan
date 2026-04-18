// 保育経営診断ツール — マネジメント共有デッキ生成
// Run: node generate-deck.js → KatagrMa_Shindan_Report.pptx

const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const {
  FaBullseye, FaBrain, FaCogs, FaChartBar, FaRocket, FaCheckCircle
} = require("react-icons/fa");

// ===== Brand =====
const LIME = "A5B830";
const GOLD = "B08D57";
const DARK = "2A2A2A";
const OFFWHITE = "F8F7F2";
const LIGHTLIME = "F0F3D8";
const ACCENT = "4A7C59";
const RED = "CC3333";
const ORANGE = "E8740C";
const GRAY = "888888";
const SUB_GRAY = "606060";
const LIGHTGRAY = "E5E5E5";
const TEXT = "333333";
const WHITE = "FFFFFF";

const FONT_JP = "Yu Gothic";
const FONT_ACCENT = "Helvetica Neue";

async function icon(IconComp, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComp, { color: "#" + color, size: String(size) })
  );
  const pngBuffer = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + pngBuffer.toString("base64");
}

// Reusable shadow factory (pptxgenjs mutates objects — need fresh each time)
const shadow = () => ({ type: "outer", blur: 8, offset: 2, color: "000000", opacity: 0.10, angle: 135 });

async function main() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.3" x 7.5"
  pres.author = "KatagrMa Beyond SaaS Team";
  pres.title = "保育経営診断ツール プロトタイプ完成報告";

  // Preload icons
  const [icBullseye, icBrain, icCogs, icChart, icRocket, icCheck] = await Promise.all([
    icon(FaBullseye, LIME),
    icon(FaBrain, LIME),
    icon(FaCogs, LIME),
    icon(FaChartBar, LIME),
    icon(FaRocket, LIME),
    icon(FaCheckCircle, LIME)
  ]);

  // ============================================================
  // Slide 1: Title (dark)
  // ============================================================
  const s1 = pres.addSlide();
  s1.background = { color: DARK };

  // Lime corner accent
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.3, h: 7.5, fill: { color: LIME }, line: { type: "none" } });

  s1.addText("KATAGRMA × AI", {
    x: 0.9, y: 1.6, w: 12, h: 0.4,
    fontSize: 13, fontFace: FONT_ACCENT, color: LIME, bold: true, charSpacing: 8, margin: 0
  });

  s1.addText("保育経営診断ツール", {
    x: 0.9, y: 2.3, w: 12, h: 1.4,
    fontSize: 60, fontFace: FONT_JP, color: WHITE, bold: true, margin: 0
  });

  s1.addText([
    { text: "プロトタイプ", options: { color: LIME } },
    { text: "完成報告", options: { color: WHITE } }
  ], {
    x: 0.9, y: 3.8, w: 12, h: 0.8,
    fontSize: 28, fontFace: FONT_JP, bold: true, margin: 0
  });

  s1.addShape(pres.shapes.LINE, { x: 0.9, y: 4.85, w: 1.8, h: 0, line: { color: LIME, width: 3 } });

  s1.addText("Beyond SaaS 戦略 Step 1 「まず懐に入る」を実現する入口ツール", {
    x: 0.9, y: 5.0, w: 12, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: OFFWHITE, bold: true, margin: 0
  });

  s1.addText("AI分析による保育施設経営診断 / MSO型マネジメントパートナーへの導線", {
    x: 0.9, y: 5.5, w: 12, h: 0.4,
    fontSize: 13, fontFace: FONT_JP, color: "BBBBBB", margin: 0
  });

  s1.addText("2026年4月  |  KatagrMa Beyond SaaS Team", {
    x: 0.9, y: 6.9, w: 12, h: 0.3,
    fontSize: 10, fontFace: FONT_ACCENT, color: "BBBBBB", charSpacing: 2, margin: 0
  });

  // ============================================================
  // Slide 2: 戦略的位置付け
  // ============================================================
  const s2 = pres.addSlide();
  s2.background = { color: OFFWHITE };

  // Top accent
  s2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 13.3, h: 0.3, fill: { color: LIME }, line: { type: "none" } });

  // Icon
  s2.addImage({ data: icBullseye, x: 0.8, y: 0.7, w: 0.55, h: 0.55 });

  s2.addText("戦略的位置付け", {
    x: 1.5, y: 0.7, w: 10, h: 0.6,
    fontSize: 32, fontFace: FONT_JP, color: DARK, bold: true, margin: 0
  });

  s2.addText("キャリアパス診断の成功パターンを「保育経営全般」へ横展開", {
    x: 1.5, y: 1.3, w: 11, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: GRAY, margin: 0
  });

  // Funnel steps
  const funnelSteps = [
    { num: "1", title: "無償診断", sub: "月間100件", color: LIME, w: 3.1 },
    { num: "2", title: "結果フィードバック", sub: "100%メール取得", color: GOLD, w: 3.0 },
    { num: "3", title: "PoC提案", sub: "目標CVR 20%", color: ORANGE, w: 2.7 },
    { num: "4", title: "MSO本契約", sub: "PoC移行率 30%", color: RED, w: 2.3 }
  ];
  let fx = 0.9;
  funnelSteps.forEach((s, i) => {
    s2.addShape(pres.shapes.RECTANGLE, {
      x: fx, y: 2.5, w: s.w, h: 2.2,
      fill: { color: WHITE }, line: { color: LIGHTGRAY, width: 1 }, shadow: shadow()
    });
    // Top color stripe
    s2.addShape(pres.shapes.RECTANGLE, {
      x: fx, y: 2.5, w: s.w, h: 0.08,
      fill: { color: s.color }, line: { type: "none" }
    });
    s2.addShape(pres.shapes.OVAL, {
      x: fx + 0.25, y: 2.8, w: 0.5, h: 0.5,
      fill: { color: s.color }, line: { type: "none" }
    });
    s2.addText(s.num, {
      x: fx + 0.25, y: 2.8, w: 0.5, h: 0.5,
      fontSize: 18, fontFace: FONT_ACCENT, color: WHITE, bold: true, align: "center", valign: "middle", margin: 0
    });
    s2.addText(s.title, {
      x: fx + 0.25, y: 3.4, w: s.w - 0.5, h: 0.5,
      fontSize: 18, fontFace: FONT_JP, color: DARK, bold: true, margin: 0
    });
    s2.addText(s.sub, {
      x: fx + 0.25, y: 3.95, w: s.w - 0.5, h: 0.4,
      fontSize: 11, fontFace: FONT_JP, color: s.color, bold: true, margin: 0
    });
    // Arrow (except after last)
    if (i < funnelSteps.length - 1) {
      s2.addText("▶", {
        x: fx + s.w, y: 3.4, w: 0.3, h: 0.4,
        fontSize: 16, color: GRAY, align: "center", valign: "middle", margin: 0
      });
    }
    fx += s.w + 0.3;
  });

  // Insight box at bottom
  s2.addShape(pres.shapes.RECTANGLE, {
    x: 0.9, y: 5.2, w: 11.5, h: 1.7,
    fill: { color: LIGHTLIME }, line: { type: "none" }
  });
  s2.addShape(pres.shapes.RECTANGLE, {
    x: 0.9, y: 5.2, w: 0.08, h: 1.7,
    fill: { color: LIME }, line: { type: "none" }
  });
  s2.addText("なぜ診断ツールか — Insight", {
    x: 1.15, y: 5.3, w: 11, h: 0.35,
    fontSize: 14, fontFace: FONT_JP, color: ACCENT, bold: true, charSpacing: 1, margin: 0
  });
  s2.addText([
    { text: "「キャリアパス診断」（Googleフォーム）をウェビナーオプションで提供したところ、", options: { breakLine: true } },
    { text: "大量の申込を獲得し、結果フィードバックを通じて継続的なアポ取得に成功。", options: { breakLine: true } },
    { text: "この成功パターンを経営診断に横展開し、AIで高度化することで", options: { color: DARK, bold: true, breakLine: true } },
    { text: "船井総研的プライムポジション（よろず相談の受け皿）の確立を目指す。", options: { color: DARK, bold: true } }
  ], {
    x: 1.15, y: 5.65, w: 11.1, h: 1.2,
    fontSize: 12, fontFace: FONT_JP, color: TEXT, margin: 0
  });

  // Footer
  s2.addText("KatagrMa 保育経営診断ツール  |  2026.04", {
    x: 0.9, y: 7.15, w: 11.5, h: 0.3,
    fontSize: 9, fontFace: FONT_ACCENT, color: GRAY, align: "right", margin: 0
  });

  // ============================================================
  // Slide 3: 診断フロー
  // ============================================================
  const s3 = pres.addSlide();
  s3.background = { color: OFFWHITE };
  s3.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 13.3, h: 0.3, fill: { color: LIME }, line: { type: "none" } });

  s3.addImage({ data: icChart, x: 0.8, y: 0.7, w: 0.55, h: 0.55 });
  s3.addText("ユーザー体験 — 5〜10分で完結する診断フロー", {
    x: 1.5, y: 0.7, w: 11, h: 0.6,
    fontSize: 28, fontFace: FONT_JP, color: DARK, bold: true, margin: 0
  });
  s3.addText("モバイルファースト設計 / プログレスバー表示 / 質問は6カテゴリ×5問の計30問", {
    x: 1.5, y: 1.3, w: 11, h: 0.4,
    fontSize: 13, fontFace: FONT_JP, color: GRAY, margin: 0
  });

  // Flow - 6 steps
  const flow = [
    { no: "01", title: "ランディング", desc: "「5分で診断」で誘導" },
    { no: "02", title: "基本情報入力", desc: "法人名・メール\n電話番号など" },
    { no: "03", title: "診断30問", desc: "6カテゴリ×5問\n5段階スケール" },
    { no: "04", title: "AI分析", desc: "Claude Opus 4.7が\n約10〜15秒で生成" },
    { no: "05", title: "結果レポート", desc: "レーダーチャート\n＋改善提案" },
    { no: "06", title: "個別相談申込", desc: "詳細計画の\n無料提案へ" }
  ];
  const boxW = 1.85, boxH = 2.5, gapX = 0.10, startX = 0.7;
  flow.forEach((f, i) => {
    const x = startX + i * (boxW + gapX);
    s3.addShape(pres.shapes.RECTANGLE, {
      x, y: 2.6, w: boxW, h: boxH,
      fill: { color: WHITE }, line: { color: LIGHTGRAY, width: 1 }, shadow: shadow()
    });
    s3.addShape(pres.shapes.RECTANGLE, {
      x, y: 2.6, w: boxW, h: 0.08,
      fill: { color: LIME }, line: { type: "none" }
    });
    s3.addText(f.no, {
      x: x + 0.15, y: 2.8, w: boxW - 0.3, h: 0.5,
      fontSize: 24, fontFace: FONT_ACCENT, color: LIME, bold: true, margin: 0
    });
    s3.addText(f.title, {
      x: x + 0.15, y: 3.4, w: boxW - 0.3, h: 0.5,
      fontSize: 14, fontFace: FONT_JP, color: DARK, bold: true, margin: 0
    });
    s3.addShape(pres.shapes.LINE, {
      x: x + 0.15, y: 3.95, w: 0.5, h: 0, line: { color: GOLD, width: 2 }
    });
    s3.addText(f.desc, {
      x: x + 0.15, y: 4.05, w: boxW - 0.3, h: 0.9,
      fontSize: 10, fontFace: FONT_JP, color: TEXT, margin: 0
    });
  });

  // Bottom metric strip
  s3.addShape(pres.shapes.RECTANGLE, {
    x: 0.7, y: 5.6, w: 11.9, h: 1.3,
    fill: { color: DARK }, line: { type: "none" }
  });
  const metrics = [
    { v: "100%", l: "メールアドレス取得率" },
    { v: "6カテゴリ", l: "経営機能網羅" },
    { v: "30問", l: "回答負荷を最小化" },
    { v: "10秒", l: "AI分析完了時間" }
  ];
  metrics.forEach((m, i) => {
    const x = 0.9 + i * 3.0;
    s3.addText(m.v, {
      x, y: 5.75, w: 2.8, h: 0.6,
      fontSize: 28, fontFace: FONT_ACCENT, color: LIME, bold: true, align: "center", margin: 0
    });
    s3.addText(m.l, {
      x, y: 6.35, w: 2.8, h: 0.4,
      fontSize: 10, fontFace: FONT_JP, color: WHITE, align: "center", margin: 0
    });
  });

  s3.addText("KatagrMa 保育経営診断ツール  |  2026.04", {
    x: 0.9, y: 7.15, w: 11.5, h: 0.3,
    fontSize: 9, fontFace: FONT_ACCENT, color: GRAY, align: "right", margin: 0
  });

  // ============================================================
  // Slide 4: AI分析エンジン
  // ============================================================
  const s4 = pres.addSlide();
  s4.background = { color: OFFWHITE };
  s4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 13.3, h: 0.3, fill: { color: LIME }, line: { type: "none" } });

  s4.addImage({ data: icBrain, x: 0.8, y: 0.7, w: 0.55, h: 0.55 });
  s4.addText("AI分析エンジン", {
    x: 1.5, y: 0.7, w: 11, h: 0.6,
    fontSize: 28, fontFace: FONT_JP, color: DARK, bold: true, margin: 0
  });
  s4.addText("最新AIで保育業界の専門用語・制度・加算に踏み込んだ所見と改善提案を自動生成", {
    x: 1.5, y: 1.3, w: 11, h: 0.4,
    fontSize: 13, fontFace: FONT_JP, color: GRAY, margin: 0
  });

  // Left panel - AI tech specs
  s4.addShape(pres.shapes.RECTANGLE, {
    x: 0.7, y: 2.1, w: 5.3, h: 4.4,
    fill: { color: DARK }, line: { type: "none" }
  });
  s4.addText("Claude Opus 4.7", {
    x: 1.0, y: 2.3, w: 5, h: 0.6,
    fontSize: 24, fontFace: FONT_ACCENT, color: LIME, bold: true, margin: 0
  });
  s4.addText("Anthropic 最新・最高性能モデル", {
    x: 1.0, y: 2.85, w: 5, h: 0.4,
    fontSize: 11, fontFace: FONT_JP, color: OFFWHITE, margin: 0
  });

  const aiFeatures = [
    { k: "適応的思考", v: "複雑な診断を自動で熟考" },
    { k: "構造化出力", v: "JSON Schema完全準拠で確実" },
    { k: "プロンプトキャッシュ", v: "コスト最適化（最大90%削減）" },
    { k: "ストリーミング", v: "タイムアウト耐性確保" }
  ];
  aiFeatures.forEach((f, i) => {
    const y = 3.5 + i * 0.75;
    s4.addShape(pres.shapes.OVAL, {
      x: 1.0, y: y + 0.1, w: 0.18, h: 0.18,
      fill: { color: LIME }, line: { type: "none" }
    });
    s4.addText(f.k, {
      x: 1.3, y, w: 4.5, h: 0.35,
      fontSize: 13, fontFace: FONT_JP, color: WHITE, bold: true, margin: 0
    });
    s4.addText(f.v, {
      x: 1.3, y: y + 0.35, w: 4.5, h: 0.35,
      fontSize: 10, fontFace: FONT_JP, color: GRAY, margin: 0
    });
  });

  // Right panel - 6 categories mapped to MSO
  s4.addText("6カテゴリ × MSO機能マッピング", {
    x: 6.4, y: 2.1, w: 6.3, h: 0.4,
    fontSize: 16, fontFace: FONT_JP, color: DARK, bold: true, margin: 0
  });

  const categories = [
    { name: "人財育成・評価", mso: "Think", color: LIME },
    { name: "採用・人員確保", mso: "Attract", color: GOLD },
    { name: "労務・制度対応", mso: "Run", color: ORANGE },
    { name: "組織マネジメント", mso: "Think", color: LIME },
    { name: "財務・経営戦略", mso: "Run", color: ORANGE },
    { name: "保護者・地域連携", mso: "2C接続", color: ACCENT }
  ];
  categories.forEach((c, i) => {
    const y = 2.6 + i * 0.63;
    s4.addShape(pres.shapes.RECTANGLE, {
      x: 6.4, y, w: 6.3, h: 0.55,
      fill: { color: WHITE }, line: { color: LIGHTGRAY, width: 1 }
    });
    s4.addShape(pres.shapes.RECTANGLE, {
      x: 6.4, y, w: 0.1, h: 0.55,
      fill: { color: c.color }, line: { type: "none" }
    });
    // Colored circle marker (replaces emoji)
    s4.addShape(pres.shapes.OVAL, {
      x: 6.65, y: y + 0.18, w: 0.2, h: 0.2,
      fill: { color: c.color }, line: { type: "none" }
    });
    s4.addText(c.name, {
      x: 6.95, y, w: 3.6, h: 0.55,
      fontSize: 13, fontFace: FONT_JP, color: DARK, bold: true, valign: "middle", margin: 0
    });
    // MSO badge
    s4.addShape(pres.shapes.RECTANGLE, {
      x: 10.8, y: y + 0.1, w: 1.7, h: 0.35,
      fill: { color: c.color }, line: { type: "none" }
    });
    s4.addText("MSO " + c.mso, {
      x: 10.8, y: y + 0.1, w: 1.7, h: 0.35,
      fontSize: 10, fontFace: FONT_ACCENT, color: WHITE, bold: true, align: "center", valign: "middle", margin: 0
    });
  });

  s4.addText("KatagrMa 保育経営診断ツール  |  2026.04", {
    x: 0.9, y: 7.15, w: 11.5, h: 0.3,
    fontSize: 9, fontFace: FONT_ACCENT, color: GRAY, align: "right", margin: 0
  });

  // ============================================================
  // Slide 5: リード取得パイプライン
  // ============================================================
  const s5 = pres.addSlide();
  s5.background = { color: OFFWHITE };
  s5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 13.3, h: 0.3, fill: { color: LIME }, line: { type: "none" } });

  s5.addImage({ data: icCogs, x: 0.8, y: 0.7, w: 0.55, h: 0.55 });
  s5.addText("リード取得から商談化までを完全自動化", {
    x: 1.5, y: 0.7, w: 11, h: 0.6,
    fontSize: 28, fontFace: FONT_JP, color: DARK, bold: true, margin: 0
  });
  s5.addText("診断完了の瞬間に、データ保存・CRM連携・メール配信・営業通知まで同時実行", {
    x: 1.5, y: 1.3, w: 11, h: 0.4,
    fontSize: 13, fontFace: FONT_JP, color: GRAY, margin: 0
  });

  // Pipeline (vertical split: diagnosis completion → fan out to 4 automated actions)
  s5.addShape(pres.shapes.RECTANGLE, {
    x: 0.7, y: 2.2, w: 3.2, h: 4.6,
    fill: { color: DARK }, line: { type: "none" }
  });
  s5.addText("診断完了", {
    x: 0.7, y: 2.5, w: 3.2, h: 0.7,
    fontSize: 32, fontFace: FONT_JP, color: LIME, bold: true, align: "center", margin: 0
  });
  s5.addText("トリガーイベント", {
    x: 0.7, y: 3.2, w: 3.2, h: 0.35,
    fontSize: 11, fontFace: FONT_JP, color: "BBBBBB", align: "center", charSpacing: 2, margin: 0
  });
  s5.addShape(pres.shapes.LINE, {
    x: 1.7, y: 3.75, w: 1.2, h: 0, line: { color: LIME, width: 2 }
  });

  // Captured data list inside trigger panel
  s5.addText("この瞬間にサーバーが受け取るデータ", {
    x: 0.95, y: 4.0, w: 2.8, h: 0.3,
    fontSize: 10, fontFace: FONT_JP, color: OFFWHITE, margin: 0
  });
  const captured = [
    "30問の回答スコア",
    "基本情報（メール必須）",
    "AI所見・改善提案",
    "UTM流入パラメータ"
  ];
  captured.forEach((c, i) => {
    const y = 4.35 + i * 0.45;
    s5.addShape(pres.shapes.OVAL, {
      x: 0.95, y: y + 0.08, w: 0.12, h: 0.12,
      fill: { color: LIME }, line: { type: "none" }
    });
    s5.addText(c, {
      x: 1.15, y, w: 2.6, h: 0.35,
      fontSize: 12, fontFace: FONT_JP, color: WHITE, valign: "middle", margin: 0
    });
  });

  // Arrow from trigger
  s5.addText("▶", {
    x: 4.0, y: 4.2, w: 0.5, h: 0.5,
    fontSize: 28, color: LIME, align: "center", valign: "middle", margin: 0
  });

  // 4 automated actions (2x2 grid)
  const actions = [
    { num: "01", title: "SQLite 永続化", desc: "全診断データを保存（PII含む22項目）", color: LIME },
    { num: "02", title: "Zoho CRM 自動登録", desc: "Webhook経由でリード化・UTM連動", color: GOLD },
    { num: "03", title: "サンキューメール", desc: "Resend経由で即時送信（PDF添付対応）", color: ACCENT },
    { num: "04", title: "営業チーム通知", desc: "相談申込時にBCC通知で即応体制", color: RED }
  ];
  const ax = 4.7, ay = 2.2, aw = 4.0, ah = 2.15, agap = 0.3;
  actions.forEach((a, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = ax + col * (aw + agap);
    const y = ay + row * (ah + agap);
    s5.addShape(pres.shapes.RECTANGLE, {
      x, y, w: aw, h: ah,
      fill: { color: WHITE }, line: { color: LIGHTGRAY, width: 1 }, shadow: shadow()
    });
    s5.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.1, h: ah, fill: { color: a.color }, line: { type: "none" }
    });
    s5.addText(a.num, {
      x: x + 0.25, y: y + 0.2, w: 1, h: 0.4,
      fontSize: 18, fontFace: FONT_ACCENT, color: a.color, bold: true, margin: 0
    });
    s5.addText(a.title, {
      x: x + 0.25, y: y + 0.7, w: aw - 0.4, h: 0.5,
      fontSize: 16, fontFace: FONT_JP, color: DARK, bold: true, margin: 0
    });
    s5.addText(a.desc, {
      x: x + 0.25, y: y + 1.25, w: aw - 0.4, h: 0.8,
      fontSize: 11, fontFace: FONT_JP, color: TEXT, margin: 0
    });
  });

  // Footer note
  s5.addShape(pres.shapes.RECTANGLE, {
    x: 0.7, y: 6.85, w: 11.9, h: 0.35,
    fill: { color: LIGHTLIME }, line: { type: "none" }
  });
  s5.addText("UTMパラメータ自動捕捉  /  Basic Auth保護された管理画面で進捗モニタリング  /  Zoho Lead ID 連携済", {
    x: 0.7, y: 6.85, w: 11.9, h: 0.35,
    fontSize: 10, fontFace: FONT_JP, color: ACCENT, bold: true, align: "center", valign: "middle", margin: 0
  });

  s5.addText("KatagrMa 保育経営診断ツール  |  2026.04", {
    x: 0.9, y: 7.25, w: 11.5, h: 0.2,
    fontSize: 9, fontFace: FONT_ACCENT, color: GRAY, align: "right", margin: 0
  });

  // ============================================================
  // Slide 6: 管理画面ダッシュボード
  // ============================================================
  const s6 = pres.addSlide();
  s6.background = { color: OFFWHITE };
  s6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 13.3, h: 0.3, fill: { color: LIME }, line: { type: "none" } });

  s6.addImage({ data: icChart, x: 0.8, y: 0.7, w: 0.55, h: 0.55 });
  s6.addText("管理画面ダッシュボード", {
    x: 1.5, y: 0.7, w: 11, h: 0.6,
    fontSize: 28, fontFace: FONT_JP, color: DARK, bold: true, margin: 0
  });
  s6.addText("Basic Auth保護。統計サマリ・診断一覧・詳細ビューで全件を一元管理", {
    x: 1.5, y: 1.3, w: 11, h: 0.4,
    fontSize: 13, fontFace: FONT_JP, color: GRAY, margin: 0
  });

  // Mock dashboard
  s6.addShape(pres.shapes.RECTANGLE, {
    x: 0.7, y: 2.1, w: 11.9, h: 4.8,
    fill: { color: WHITE }, line: { color: LIGHTGRAY, width: 1 }, shadow: shadow()
  });
  // Browser-like top bar
  s6.addShape(pres.shapes.RECTANGLE, {
    x: 0.7, y: 2.1, w: 11.9, h: 0.4,
    fill: { color: "F0F0F0" }, line: { type: "none" }
  });
  s6.addShape(pres.shapes.OVAL, { x: 0.85, y: 2.22, w: 0.15, h: 0.15, fill: { color: "FF5F56" }, line: { type: "none" } });
  s6.addShape(pres.shapes.OVAL, { x: 1.08, y: 2.22, w: 0.15, h: 0.15, fill: { color: "FFBD2E" }, line: { type: "none" } });
  s6.addShape(pres.shapes.OVAL, { x: 1.31, y: 2.22, w: 0.15, h: 0.15, fill: { color: "27C93F" }, line: { type: "none" } });
  s6.addText("admin.katagrma.jp/shindan", {
    x: 1.6, y: 2.1, w: 5, h: 0.4,
    fontSize: 10, fontFace: FONT_ACCENT, color: GRAY, valign: "middle", margin: 0
  });

  // KPI cards
  s6.addText("診断結果 管理ダッシュボード", {
    x: 1.0, y: 2.7, w: 6, h: 0.4,
    fontSize: 16, fontFace: FONT_JP, color: DARK, bold: true, margin: 0
  });
  s6.addText("（表示は参考イメージ）", {
    x: 7.0, y: 2.7, w: 5.4, h: 0.4,
    fontSize: 10, fontFace: FONT_JP, color: GRAY, align: "right", margin: 0
  });

  const kpis = [
    { v: "247", l: "総診断数", c: DARK },
    { v: "52", l: "相談申込数", c: LIME },
    { v: "21.1%", l: "コンバージョン率", c: GOLD },
    { v: "58", l: "平均スコア", c: DARK }
  ];
  kpis.forEach((k, i) => {
    const x = 1.0 + i * 2.85;
    s6.addShape(pres.shapes.RECTANGLE, {
      x, y: 3.25, w: 2.65, h: 1.2,
      fill: { color: OFFWHITE }, line: { type: "none" }
    });
    s6.addText(k.l, {
      x: x + 0.15, y: 3.35, w: 2.4, h: 0.3,
      fontSize: 10, fontFace: FONT_JP, color: GRAY, margin: 0
    });
    s6.addText(k.v, {
      x: x + 0.15, y: 3.7, w: 2.4, h: 0.7,
      fontSize: 36, fontFace: FONT_ACCENT, color: k.c, bold: true, margin: 0
    });
  });

  // Mock table
  const tableHeaders = ["日時", "法人名", "施設数", "スコア", "ランク", "相談"];
  const tableRows = [
    ["04-19 12:34", "社会福祉法人 A会", "3", "54", "B", "申込"],
    ["04-19 11:02", "学校法人 B学園", "8", "72", "A", "-"],
    ["04-18 18:44", "株式会社 C保育", "2", "45", "B", "申込"],
    ["04-18 15:11", "NPO法人 D", "1", "38", "C", "-"]
  ];
  // Header
  s6.addShape(pres.shapes.RECTANGLE, {
    x: 1.0, y: 4.75, w: 11.35, h: 0.3,
    fill: { color: DARK }, line: { type: "none" }
  });
  tableHeaders.forEach((h, i) => {
    const widths = [2.2, 3.5, 1.3, 1.3, 1.4, 1.65];
    const x = 1.0 + widths.slice(0, i).reduce((s, w) => s + w, 0);
    s6.addText(h, {
      x: x + 0.15, y: 4.75, w: widths[i], h: 0.3,
      fontSize: 10, fontFace: FONT_ACCENT, color: WHITE, bold: true, valign: "middle", margin: 0
    });
  });
  tableRows.forEach((row, ri) => {
    const y = 5.1 + ri * 0.35;
    if (ri % 2 === 0) {
      s6.addShape(pres.shapes.RECTANGLE, {
        x: 1.0, y, w: 11.35, h: 0.35,
        fill: { color: "FAFAFA" }, line: { type: "none" }
      });
    }
    row.forEach((cell, i) => {
      const widths = [2.2, 3.5, 1.3, 1.3, 1.4, 1.65];
      const x = 1.0 + widths.slice(0, i).reduce((s, w) => s + w, 0);
      s6.addText(cell, {
        x: x + 0.15, y, w: widths[i], h: 0.35,
        fontSize: 10, fontFace: FONT_JP, color: TEXT, valign: "middle", margin: 0
      });
    });
  });

  s6.addText("KatagrMa 保育経営診断ツール  |  2026.04", {
    x: 0.9, y: 7.15, w: 11.5, h: 0.3,
    fontSize: 9, fontFace: FONT_ACCENT, color: GRAY, align: "right", margin: 0
  });

  // ============================================================
  // Slide 7: KPI目標 & 実装状況
  // ============================================================
  const s7 = pres.addSlide();
  s7.background = { color: OFFWHITE };
  s7.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 13.3, h: 0.3, fill: { color: LIME }, line: { type: "none" } });

  s7.addImage({ data: icCheck, x: 0.8, y: 0.7, w: 0.55, h: 0.55 });
  s7.addText("KPI目標 & 実装状況", {
    x: 1.5, y: 0.7, w: 11, h: 0.6,
    fontSize: 28, fontFace: FONT_JP, color: DARK, bold: true, margin: 0
  });
  s7.addText("Phase 1 / Phase 2 の実装は完了。あとは本番デプロイと質問文最終化のみ", {
    x: 1.5, y: 1.3, w: 11, h: 0.4,
    fontSize: 13, fontFace: FONT_JP, color: GRAY, margin: 0
  });

  // Top row: KPI targets (4 big stats)
  s7.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: 2.08, w: 0.08, h: 0.32, fill: { color: LIME }, line: { type: "none" } });
  s7.addText("目標KPI", {
    x: 0.85, y: 2.0, w: 6, h: 0.4,
    fontSize: 15, fontFace: FONT_JP, color: DARK, bold: true, charSpacing: 2, margin: 0
  });

  const kpiTargets = [
    { v: "100+", l: "月間診断完了数" },
    { v: "100%", l: "メールアドレス取得率" },
    { v: "20%+", l: "相談申込CVR" },
    { v: "30%+", l: "PoC本契約移行率" }
  ];
  kpiTargets.forEach((k, i) => {
    const x = 0.7 + i * 3.0;
    s7.addShape(pres.shapes.RECTANGLE, {
      x, y: 2.45, w: 2.85, h: 1.6,
      fill: { color: WHITE }, line: { color: LIGHTGRAY, width: 1 }, shadow: shadow()
    });
    s7.addShape(pres.shapes.RECTANGLE, {
      x, y: 2.45, w: 2.85, h: 0.08,
      fill: { color: LIME }, line: { type: "none" }
    });
    s7.addText(k.v, {
      x: x + 0.1, y: 2.65, w: 2.65, h: 0.9,
      fontSize: 44, fontFace: FONT_ACCENT, color: DARK, bold: true, align: "center", margin: 0
    });
    s7.addText(k.l, {
      x: x + 0.1, y: 3.55, w: 2.65, h: 0.4,
      fontSize: 11, fontFace: FONT_JP, color: GRAY, align: "center", margin: 0
    });
  });

  // Bottom: implementation status (phase cards)
  s7.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: 4.43, w: 0.08, h: 0.32, fill: { color: LIME }, line: { type: "none" } });
  s7.addText("実装状況", {
    x: 0.85, y: 4.35, w: 6, h: 0.4,
    fontSize: 15, fontFace: FONT_JP, color: DARK, bold: true, charSpacing: 2, margin: 0
  });

  const phases = [
    {
      title: "Phase 1 — MVP",
      status: "完了",
      color: LIME,
      items: "診断UI（6画面30問）/ スコアリング / Claude Opus 4.7 AI分析 / レーダーチャート結果レポート / モバイル対応"
    },
    {
      title: "Phase 2 — リード最適化",
      status: "完了",
      color: LIME,
      items: "SQLite永続化 / Zoho CRM Webhook / サンキューメール / 営業通知 / 管理画面ダッシュボード"
    },
    {
      title: "Phase 3 — 本番化",
      status: "着手予定",
      color: GOLD,
      items: "質問30問の最終化 / 535施設ベンチマーク算出 / 本番デプロイ / ウェビナー連携 / A/Bテスト基盤"
    }
  ];
  phases.forEach((p, i) => {
    const x = 0.7 + i * 4.0;
    s7.addShape(pres.shapes.RECTANGLE, {
      x, y: 4.8, w: 3.85, h: 2.0,
      fill: { color: WHITE }, line: { color: LIGHTGRAY, width: 1 }, shadow: shadow()
    });
    s7.addShape(pres.shapes.RECTANGLE, {
      x, y: 4.8, w: 0.12, h: 2.0,
      fill: { color: p.color }, line: { type: "none" }
    });
    s7.addText(p.title, {
      x: x + 0.3, y: 4.95, w: 2.5, h: 0.4,
      fontSize: 14, fontFace: FONT_JP, color: DARK, bold: true, margin: 0
    });
    // Status badge
    s7.addShape(pres.shapes.RECTANGLE, {
      x: x + 2.85, y: 5.0, w: 0.9, h: 0.3,
      fill: { color: p.color }, line: { type: "none" }
    });
    s7.addText(p.status, {
      x: x + 2.85, y: 5.0, w: 0.9, h: 0.3,
      fontSize: 10, fontFace: FONT_JP, color: WHITE, bold: true, align: "center", valign: "middle", margin: 0
    });
    s7.addText(p.items, {
      x: x + 0.3, y: 5.45, w: 3.4, h: 1.3,
      fontSize: 10, fontFace: FONT_JP, color: TEXT, margin: 0
    });
  });

  s7.addText("KatagrMa 保育経営診断ツール  |  2026.04", {
    x: 0.9, y: 7.15, w: 11.5, h: 0.3,
    fontSize: 9, fontFace: FONT_ACCENT, color: GRAY, align: "right", margin: 0
  });

  // ============================================================
  // Slide 8: Next Steps (dark)
  // ============================================================
  const s8 = pres.addSlide();
  s8.background = { color: DARK };
  s8.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.3, h: 7.5, fill: { color: LIME }, line: { type: "none" } });

  s8.addImage({ data: icRocket, x: 0.9, y: 0.7, w: 0.55, h: 0.55 });
  s8.addText("次のステップ", {
    x: 1.6, y: 0.7, w: 11, h: 0.6,
    fontSize: 32, fontFace: FONT_JP, color: WHITE, bold: true, margin: 0
  });
  s8.addText("MVPは完成。残タスクを順次消化することで、ウェビナー集客テストへ展開可能", {
    x: 1.6, y: 1.3, w: 11, h: 0.4,
    fontSize: 13, fontFace: FONT_JP, color: OFFWHITE, margin: 0
  });

  const nextSteps = [
    {
      num: "01",
      title: "質問30問の最終化",
      owner: "大嶽CEO監修",
      timeline: "1週間",
      desc: "プレースホルダーを保育コンサル経験ベースで差し替え。JSONから読み込む構造のため、コード変更不要。"
    },
    {
      num: "02",
      title: "535施設ベンチマーク算出",
      owner: "データチーム",
      timeline: "2週間",
      desc: "現在は推定値（業界平均52/38/45/48/42/55）。実データから算出して差し替え、診断精度を向上。"
    },
    {
      num: "03",
      title: "本番デプロイ & ドメイン",
      owner: "エンジニア",
      timeline: "3日",
      desc: "shindan.katagrma.jp へ公開。ウェビナー LP / メルマガから誘導開始。"
    }
  ];
  nextSteps.forEach((n, i) => {
    const x = 0.9 + i * 4.1;
    // Card
    s8.addShape(pres.shapes.RECTANGLE, {
      x, y: 2.3, w: 3.9, h: 4.3,
      fill: { color: "3A3A3A" }, line: { color: "4A4A4A", width: 1 }
    });
    // Big number
    s8.addText(n.num, {
      x: x + 0.3, y: 2.5, w: 2, h: 1.0,
      fontSize: 72, fontFace: FONT_ACCENT, color: LIME, bold: true, margin: 0
    });
    // Title
    s8.addText(n.title, {
      x: x + 0.3, y: 3.6, w: 3.4, h: 0.6,
      fontSize: 18, fontFace: FONT_JP, color: WHITE, bold: true, margin: 0
    });
    // Divider
    s8.addShape(pres.shapes.LINE, {
      x: x + 0.3, y: 4.3, w: 0.6, h: 0, line: { color: LIME, width: 2 }
    });
    // Meta
    s8.addText([
      { text: "担当: ", options: { color: GRAY } },
      { text: n.owner, options: { color: OFFWHITE, bold: true, breakLine: true } },
      { text: "期間: ", options: { color: GRAY } },
      { text: n.timeline, options: { color: LIME, bold: true } }
    ], {
      x: x + 0.3, y: 4.45, w: 3.4, h: 0.8,
      fontSize: 11, fontFace: FONT_JP, margin: 0
    });
    // Desc
    s8.addText(n.desc, {
      x: x + 0.3, y: 5.3, w: 3.4, h: 1.2,
      fontSize: 10, fontFace: FONT_JP, color: OFFWHITE, margin: 0
    });
  });

  // Bottom CTA
  s8.addShape(pres.shapes.RECTANGLE, {
    x: 0.9, y: 6.75, w: 11.5, h: 0.5,
    fill: { color: LIME }, line: { type: "none" }
  });
  s8.addText("着手順で進めれば、2026年5月末までにウェビナー集客テストを開始可能", {
    x: 0.9, y: 6.75, w: 11.5, h: 0.5,
    fontSize: 14, fontFace: FONT_JP, color: DARK, bold: true, align: "center", valign: "middle", margin: 0
  });

  // ============================================================
  // Write file
  // ============================================================
  await pres.writeFile({ fileName: "KatagrMa_Shindan_Report.pptx" });
  console.log("✅ Generated: KatagrMa_Shindan_Report.pptx");
}

main().catch(err => { console.error(err); process.exit(1); });
