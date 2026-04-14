# GAS認証セキュリティ設定手順

touka-tools Web App のアクセスを `@devine.co.jp` の指定メンバー3名に限定するための手順書です。

## 許可メンバー

| 氏名 | メールアドレス |
|------|---------------|
| 依田知則 | tomonori.yorita@devine.co.jp |
| 大久保直登 | naoto.okubo@devine.co.jp |
| 三原園子 | sonoko.mihara@devine.co.jp |

このリストは `Code_gs_complete.js` の定数 `ALLOWED_USERS` にハードコードしてあります。メンバー変更時はこの配列を編集して再デプロイします。

## 全体の構成

- **GAS側のゲート**: `checkAccess()` が `Session.getActiveUser().getEmail()` を allowlist と照合し、許可されていないアカウントは即エラー返却
- **Google側のゲート**: GAS Web App デプロイ時に「アクセス権」を絞ることでGoogleログイン自体を必須化
- **シート・Driveのゲート**: 共有権限を3メンバーのみに限定することで、GASを経由しない直接アクセスも遮断
- 3層の多層防御で守ります。

## 実施手順

### ステップ1: GASコードを最新版に更新

1. `C:\Users\hashi\touka-tools\copy_gas.ps1` を実行し、`Code_gs_complete.js` の内容をクリップボードにコピー
2. GASエディタを開き、既存の `Code.gs` の中身を全削除してペースト
3. 保存（Ctrl+S）

コード冒頭に以下が追加されているかを確認：
- `ALLOWED_USERS` 定数に3名のメールアドレス
- `checkAccess()` 関数
- `buildUnauthorizedResponse()` 関数
- `doGet` / `doPost` の先頭に `checkAccess()` 呼び出し

### ステップ2: GAS Web App のデプロイ設定変更

GASエディタ右上の「デプロイ」→「デプロイの管理」を開き、既存デプロイの編集（鉛筆アイコン）を選択。

以下の設定に変更：

| 項目 | 設定値 |
|------|-------|
| 種類 | ウェブアプリ |
| 実行するユーザー | **自分**（シート所有者） |
| アクセスできるユーザー | **devine.co.jp 内のユーザー**（ドメイン内のみ） |

「バージョン」を「新バージョン」にして「デプロイ」をクリック。**Web App URLは変わらない**のでHTML側の修正は不要（URLを維持したままデプロイ管理から設定変更）。

> 注意: Workspaceを使っていれば「devine.co.jp 内のユーザー」の選択肢が表示されます。表示されない場合は「Googleアカウントを持つすべてのユーザー」にして、代わりに `checkAccess()` の allowlist だけで守る形にもできます（Googleログインは必須になります）。

### ステップ3: Googleスプレッドシートの共有権限を絞る

対象3シート：

| シート名 | ID |
|---------|-----|
| メインシート（回答データ） | `1znspyaI-wj70aBkDfOPPrmYjPSSn3mFELW6VNfOSIbM` |
| 検体PpP値シート | `1y4z8dZjKuJKOS0HUxmnSSfkkvhRWkQpb--TXGm5amvs` |
| 追跡データ（HbA1c）シート | `1jr25zPPv2qCHkgXNA6oHV5eoSq0RcxpXEirWq0V5YA8` |

各シートで「共有」→「一般的なアクセス」を**「制限付き」**に変更。その上で3名を**編集者**として個別に追加。

「リンクを知っている全員」になっていたら必ず解除してください。

### ステップ4: Google Driveフォルダの共有権限を絞る

対象3フォルダ：

| フォルダ名 | ID |
|-----------|-----|
| 元PDFフォルダ | `1blx2Ia2X9blWcYAkGKPGST8MFN3VgJGL` |
| 入力済みPDFフォルダ | `1iFx7ngwNSJo80bElYcqw0dyHQyhm0N0N` |
| アンケートJSON（親フォルダ） | `1tAUwyUb9B-WH5LrW45-ox4GZFrPJWMnB` |

同様に「一般的なアクセス」を「制限付き」にし、3名を編集者として追加。子フォルダ（事前OCR、OCR照合結果）は親から継承されます。

### ステップ5: 動作確認

以下の観点でテストします。

**許可メンバーでの動作確認**
- 3名それぞれの `@devine.co.jp` アカウントでログインした状態でダッシュボードを開き、各ツール（データ閲覧・検体番号検索・統合エクスポート・追跡データ入力など）が従来通り動く

**未許可ユーザーでの遮断確認**
- Googleアカウント未ログイン状態 → Google認証画面が表示される
- `@devine.co.jp` 以外のGoogleアカウントでログイン → `unauthorized: true` のエラーレスポンスが返る
- スプレッドシート・Driveフォルダを未許可アカウントでURL直叩き → 「権限がありません」表示

**allowlist追加忘れパターンの確認**
- ドメイン内だがallowlist外の `@devine.co.jp` アカウント（仮にあれば） → `アクセス権限がありません` エラー

### ステップ6: GitHub Pagesへのpush（コード変更なし）

HTML側は今回変更していないためpush不要ですが、`Code_gs_complete.js` をリポジトリにも反映させておくとgit履歴に残ります。

```
git add Code_gs_complete.js docs/GAS認証セキュリティ設定手順.md
git commit -m "feat: GAS Web App にメンバー限定allowlist認証を追加"
git push
```

## 運用上の注意

**メンバー追加・削除時**
1. `Code_gs_complete.js` の `ALLOWED_USERS` 配列を編集
2. GASエディタにペースト→保存
3. 「デプロイ管理」で既存デプロイを編集→「新バージョン」でデプロイ
4. シート3つ・Driveフォルダ3つの共有権限も同期して追加・削除

**パスワード定数 `VIEW_PASSWORD = 'touka2026'` について**
現行の一部ツール（データ閲覧・検体番号検索・統合エクスポート・AIアシスタント）には古いパスワードゲートが残っています。allowlistによる二重防御状態なので当面は残しても害はありませんが、整理するなら以降のPRで削除しても構いません。

**Googleフォームへの影響**
`onFormSubmit` トリガーはGAS所有者権限で自動実行されるので、患者側のフォーム回答フローは**今回の変更で一切影響を受けません**。

**注意: JSONP GETの限界**
URLクエリにトークン等を載せる設計は引き続き残ります。今回はallowlistが `Session.getActiveUser()` ベースなのでURLクエリでトークンを運ぶ必要はなく、**クエリパラメータに認証情報は含まれません**（Googleログインセッションのみで判定）。この点では設計が改善されています。

## ロールバック手順

万一既存機能が壊れた場合：
1. GASエディタ「デプロイ管理」で1つ前のバージョンに戻す
2. シート・Driveの共有設定を元に戻す（必要なら）

コード変更前のバックアップは git 履歴から参照可能です。
