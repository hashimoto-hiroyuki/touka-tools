# 糖化アンケート管理システム — touka-tools

> Web管理ダッシュボード（GitHub Pages）
> 🔗 **https://hashimoto-hiroyuki.github.io/touka-tools/**

歯科医院で実施する糖化アンケートの **回収 → OCR読み取り → 照合 → データ統合** までの全プロセスを支援するWebツール群です。

---

## 目次

- [システム全体像](#システム全体像)
- [アンケートデータの入力手順（ワークフロー）](#アンケートデータの入力手順ワークフロー)
  - [Step 1. PDFファイルの2ページ化](#step-1-pdfファイルの2ページ化)
  - [Step 2. 事前AI-OCR処理](#step-2-事前ai-ocr処理)
  - [Step 3. 照合作業](#step-3-照合作業)
  - [Step 4. サンプルのデータ番号記入](#step-4-サンプルのデータ番号記入)
  - [Step 5. PpP値の入力](#step-5-ppp値の入力)
  - [Step 6. 追跡データの追加入力](#step-6-追跡データの追加入力)
  - [Step 7. 統合データのエクスポート](#step-7-統合データのエクスポート)
- [ツール一覧](#ツール一覧)
- [データストア](#データストア)
- [関連プロジェクト](#関連プロジェクト)
- [技術スタック](#技術スタック)

---

## システム全体像

```
┌──────────────────┐     ┌───────────────────┐     ┌──────────────────┐
│  touka-survey    │     │   touka-tools     │     │   Google Sheets  │
│  (Vercel)        │     │  (GitHub Pages)   │     │ + Google Drive   │
│                  │     │                   │     │                  │
│ アンケート用紙作成 │     │ 管理ダッシュボード  │     │ 共通データストア   │
│ QRコード生成      │────▶│ OCR照合・PDF紐づけ  │────▶│ 回答データ        │
│ 医療機関管理      │     │ 検査値入力         │     │ PpP値            │
│                  │     │ データエクスポート  │     │ 追跡データ        │
└──────────────────┘     └───────────────────┘     └──────────────────┘
                                  │
                          GAS Web App API
                         （バックエンド処理）
```

本リポジトリ（**touka-tools**）は、ダッシュボード上に配置された **29のWebツール** を提供し、アンケートの回収からビッグデータ構築までの全作業を効率化します。

---

## アンケートデータの入力手順（ワークフロー）

回収したアンケートから統合データシートを構築するまでの手順です。
スタッフはこの手順に沿って作業を進めます。

```
 ① スキャン    ② Drive保存    ③ 事前OCR     ④⑤ 照合       ⑥ 検体No.     ⑦ PpP値      ⑧ 追跡データ   ⑨ 統合
 ─────────── ─────────── ─────────── ─────────── ─────────── ─────────── ─────────── ───────────
 アンケート用紙  元PDFフォルダ   OCRフォルダ    スプレッド     サンプル袋     PEN/PYD      HbA1c等      CSV
 →複合機で     →2ページ化     →AI一括OCR    シートへ反映   →No.記入     →PpP算出     →追加記録    →ダウンロード
   スキャン       して保存       (json保存)     +PDF紐づけ
```

---

### Step 1. PDFファイルの2ページ化

回収したアンケート用紙（表面＝患者記入、裏面＝医師記入）をスキャンし、**表裏2ページが1セット**のPDFファイルにします。

| 作業 | 説明 | 使用アプリ |
|------|------|-----------|
| スキャン | 複合機でアンケート用紙をまとめてスキャン | 複合機 |
| PDF分割 | まとめスキャンPDFを **2ページずつ** に一括分割 | **スキャンPDF分割** |
| 結合（必要時） | 表・裏が別々のPDFファイルの場合、1つに結合 | 手動 or ツール |
| 確認 | 元PDFフォルダに表裏セットのPDFが保存されていることを確認 | **元PDF閲覧ビューア** |

> 📁 元PDFフォルダ（Google Drive）には、表と裏がセットになった1つのPDFファイルが保存されている状態にしてください。

---

### Step 2. 事前AI-OCR処理

照合作業の前に、**全PDFファイルをAIで一括OCR処理**しておきます。

| 作業 | 説明 | 使用アプリ |
|------|------|-----------|
| 一括OCR実行 | 元PDFフォルダの全PDFをGemini AIで自動OCR | **全PDF_to_OCR** |

- **なぜ事前処理が必要か？**
  照合作業と同時にOCRすると、AIによるOCR処理時間が **1件あたり数十秒** かかり、待ち時間が発生して非効率です。
  事前に一括でOCR処理しておくことで、照合時には即座に結果が表示されます。

- OCR結果は **事前OCRフォルダにJSONファイル** として一時保存されます。

---

### Step 3. 照合作業

スタッフがOCR結果を目視で確認し、読み取りミスがあれば修正します。
**QRコードチェックの有無**で作業フローが分岐します。

#### ☐ QRコードチェックなし（紙で全問回答）の場合

| 作業 | 説明 | 使用アプリ |
|------|------|-----------|
| OCR照合 | 事前OCRのJSON結果とPDF画像を並べて照合。読み取りミスは手動修正 | **☐QRコード(紙回答)\_OCR\_to\_目視チェック** |

- 照合が完了すると、**回答データとPDFファイルが紐づけ**られます。
- 回答データがスプレッドシートに追加され、JSONファイルも保存されます。

#### ☑ QRコードチェックあり（オンラインで回答済み）の場合

| 作業 | 説明 | 使用アプリ |
|------|------|-----------|
| PDF紐づけ | 患者がQRコードからオンラインで回答済み。データシートには回答が記録済みだが、**医師の記入データがない**。PDFの医療機関名・氏名・生年月日から該当データを検索し、PDFと回答データを紐づける | **☑QRコード(フォーム回答)\_回答データ\_to\_PDF** |

- データシートには回答結果がすでに記録されており、JSONファイルも作成済みです。
- 紐づけ時に、PDFから **抜歯位置・コメント**（医師記入欄）を追記します。

---

### Step 4. サンプルのデータ番号記入

検体サンプルの袋には **名前・生年月日** が記入されています。
この情報をもとにデータシートのデータNo.を検索し、サンプル袋にNo.を記入します。

| 作業 | 説明 | 使用アプリ |
|------|------|-----------|
| 検体番号検索 | 名前・生年月日であいまい検索し、データNo.を特定 | **検体番号検索** |

> これにより、**検体サンプルと回答データが紐づき**ます。

---

### Step 5. PpP値の入力

PpP値は分析完了後に数値が得られるため、照合作業とは別のタイミングで入力します。

| 作業 | 説明 | 使用アプリ |
|------|------|-----------|
| PEN/PYD入力 | 検体のPEN値・PYD値を入力すると、PpP値が自動計算される | **検体PpP値 入力・閲覧** |

- PpP値 = PEN ÷ PYD × 1000 で自動算出
- 入力データはJSONファイルにも反映・記録されます。

---

### Step 6. 追跡データの追加入力

HbA1cなど、医師が特に記録すべき追跡情報があれば追加入力します。

| 作業 | 説明 | 使用アプリ |
|------|------|-----------|
| 追跡データ入力 | HbA1c、X-CT画像URL、分析値などの数値・URLを入力 | **追跡データ入力・閲覧・編集** |

- 入力データはJSONファイルにも反映・記録されます。

---

### Step 7. 統合データのエクスポート

全ステップで蓄積されたデータを統合し、ビッグデータとして出力します。

| 作業 | 説明 | 使用アプリ |
|------|------|-----------|
| 統合エクスポート | アンケート回答 + 医師記入 + PpP値 + 追跡データを **患者1人=1行** に横展開してスプレッドシート表示、CSVダウンロード | **統合データエクスポート** |

- 3つのスプレッドシート（回答データ・PpP値・追跡データ）を患者No.をキーに結合
- CSVファイルでダウンロードし、統計ソフト等で分析可能

---

## ツール一覧

### 📄 アンケート作成・開発

| ツール | 種別 | 説明 |
|--------|------|------|
| [アンケート作成ツール](https://touka-survey.vercel.app) | Vercel | 医療機関を選択し、QRコード付きA4アンケート用紙を作成・印刷 |
| [アンケート白紙](https://hashimoto-hiroyuki.github.io/touka-tools/白紙.pdf) | Web | 医療機関名・QRコードなしの白紙アンケート用紙PDF |
| [スキャンPDF分割](https://hashimoto-hiroyuki.github.io/touka-tools/split_scan_pdf.html) | Web | まとめスキャンPDFを2ページずつに一括分割 |
| [アンケートフォーム](https://docs.google.com/forms/d/e/1FAIpQLSfK29rSSrvSjt7onYIO5gDCLDhtj776z-EhKfTxf2gUlGPBlQ/viewform) | Google | Googleフォーム（QRコードからの回答先） |

### 🔍 OCR読み取り・照合

| ツール | 種別 | 説明 |
|--------|------|------|
| [全PDF_to_OCR](https://hashimoto-hiroyuki.github.io/touka-tools/pre_ocr.html) | Web | 元PDFフォルダの全PDFをGemini AIで事前一括OCR処理 |
| [☐QRコード(紙回答)\_OCR\_to\_目視チェック](https://hashimoto-hiroyuki.github.io/touka-tools/pre_ocr_verify.html) | Web | 紙回答PDFのOCR結果を目視照合。読み取りミスを手動修正 |
| [☑QRコード(フォーム回答)\_回答データ\_to\_PDF](https://hashimoto-hiroyuki.github.io/touka-tools/link_pdf.html) | Web | フォーム回答済みデータとPDFを紐づけ。医師記入欄を追記 |
| [OCR照合（共有版）](https://hashimoto-hiroyuki.github.io/touka-tools/verify_ocr.html) | Web | OCR照合の共有版 |
| [PDF先読み照合](https://hashimoto-hiroyuki.github.io/touka-tools/prefetch_verify.html) | Web | PDFを先読みして効率的に照合 |
| [Web OCR照合](https://hashimoto-hiroyuki.github.io/touka-tools/web_ocr_verify.html) | Web | Google Drive上のPDFをブラウザで直接OCR照合（左右パネル同期） |
| [紙回答自動OCR+照合](https://hashimoto-hiroyuki.github.io/touka-tools/paper_ocr_verify.html) | Web | 紙回答の自動OCR処理と照合支援 |
| [バッチOCR照合](https://hashimoto-hiroyuki.github.io/touka-tools/batch_verify_ocr.html) | Web | 複数PDFのバッチOCR照合 |
| [PDF処理状況チェック](https://hashimoto-hiroyuki.github.io/touka-tools/check_pdf_status.html) | Web | PDFの処理状況を確認 |

### 🗄️ データ管理

| ツール | 種別 | 説明 |
|--------|------|------|
| [データ閲覧](https://hashimoto-hiroyuki.github.io/touka-tools/view_data.html) | Web | スプレッドシートのデータをWeb上で閲覧（パスワード保護） |
| [検体番号検索](https://hashimoto-hiroyuki.github.io/touka-tools/search_sample.html) | Web | 名前・生年月日であいまい検索し、データNo.を特定 |
| [PDFインデックス構築](https://hashimoto-hiroyuki.github.io/touka-tools/index_pdfs.html) | Web | Gemini AIでPDFの内容をインデックス化 |
| [PDF到着状況](https://hashimoto-hiroyuki.github.io/touka-tools/pdf_arrival_status.html) | Web | PDFファイルの到着・処理状況を確認 |
| [JSON復元](https://hashimoto-hiroyuki.github.io/touka-tools/restore_from_json.html) | Web | JSONファイルからデータを復元 |
| [統合データエクスポート](https://hashimoto-hiroyuki.github.io/touka-tools/bigdata_export.html) | Web | 3シート統合→患者1行→CSV出力 |
| [テストデータ削除](https://hashimoto-hiroyuki.github.io/touka-tools/delete_test_data.html) | Web | テスト用データの削除 |

### 🧪 検査値入力

| ツール | 種別 | 説明 |
|--------|------|------|
| [検体PpP値 入力・閲覧](https://hashimoto-hiroyuki.github.io/touka-tools/PpP_input.html) | Web | PEN/PYD値入力 → PpP値自動計算 |
| [追跡データ入力・閲覧・編集](https://hashimoto-hiroyuki.github.io/touka-tools/HbA1c_tracking_input.html) | Web | HbA1c等の追跡データ入力・閲覧・編集 |
| PpP値データ | Sheets | PpP値スプレッドシート（直接編集） |
| 追跡データ | Sheets | 追跡データスプレッドシート（直接編集） |

### 📊 データフロー・リファレンス

| ツール | 種別 | 説明 |
|--------|------|------|
| [全体データフロー](https://hashimoto-hiroyuki.github.io/touka-tools/data_flow_chart.html) | Web | システム全体のデータフロー図 |
| [作業フロー図](https://hashimoto-hiroyuki.github.io/touka-tools/workflow_chart.html) | Web | スタッフの作業フロー図 |
| [元PDF閲覧ビューア](https://hashimoto-hiroyuki.github.io/touka-tools/view_pdf.html) | Web | Google Drive上の元PDFを閲覧 |
| [ナレッジベース](https://hashimoto-hiroyuki.github.io/touka-tools/knowledge_base.html) | Web | 全ツールの概要・アルゴリズム検索 |
| [AIアシスタント](https://hashimoto-hiroyuki.github.io/touka-tools/assistant_chat.html) | Web | Gemini AIによるシステムQ&Aチャットボット |

---

## データストア

| リソース | 内容 |
|---------|------|
| **スプレッドシート（メイン）** | フォーム回答 + OCR照合データ。シート「フォームの回答 1」「医療機関リスト」 |
| **スプレッドシート（PpP値）** | 検体のPEN/PYD/PpP値 |
| **スプレッドシート（追跡データ）** | HbA1c等の追跡データ |
| **Google Drive — 元PDFフォルダ** | スキャンした元PDF（表裏2ページ/件） |
| **Google Drive — 入力済みPDFフォルダ** | 照合完了後のPDF |
| **Google Drive — JSONフォルダ** | 各患者のOCR結果・回答データJSON |

---

## 関連プロジェクト

| プロジェクト | リポジトリ | デプロイ先 | 役割 |
|-------------|-----------|-----------|------|
| **touka-survey** | [GitHub](https://github.com/hashimoto-hiroyuki/touka-survey) | [Vercel](https://touka-survey.vercel.app) | アンケート用紙作成（React + Vite） |
| **touka-tools** | [GitHub](https://github.com/hashimoto-hiroyuki/touka-tools) | [GitHub Pages](https://hashimoto-hiroyuki.github.io/touka-tools/) | 本リポジトリ。Web管理ダッシュボード |
| **test-OCR** | ローカル | ローカル実行 | スキャンPDFのOCR処理（Python） |

---

## 技術スタック

| 技術 | 用途 |
|------|------|
| HTML / CSS / JavaScript | 各WebツールのフロントエンドUI |
| Google Apps Script (GAS) | バックエンドAPI（Web App） |
| Gemini AI API | OCR読み取り・PDFインデックス・AIアシスタント |
| Google Sheets API | データの読み書き |
| Google Drive API | PDFファイルの管理 |
| GitHub Pages | 本ダッシュボードのホスティング |
| PDF.js | ブラウザ上でのPDFレンダリング |

---

## ライセンス

Private — 社内利用限定
