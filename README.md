## 🖌️ 書道教室予約システム (liff_calligraphy_class_reservation)  
LINE LIFFを活用した、シンプルかつ柔軟な予約管理システムです。  
Google Apps Script (GAS) をバックエンド、Google スプレッドシートをデータベースとして利用します。

## 🌟 主な機能
- LINE連携予約: LINEからログイン不要（LINE認証）で簡単に予約・キャンセルが可能。
- 2ヶ月ローリング予約: 当月および翌月の予約枠をリアルタイムに表示。
- 個別予約回数管理: 生徒ごとに「今月○回」「来月○回」といった予約上限を個別に設定可能。
- 自動メンテナンス機能: 毎月1日に以下の処理を自動実行。
  - 先月以前の古い予約データの削除（軽量化）。
  - 「来月の残り回数」を「今月の残り回数」へ自動スライド。
  - 来月分の回数を各生徒のデフォルト値でリセット。
- Googleカレンダー同期: 1時間毎にGoogleカレンダー同期処理（トリガー）を実行。
- LINEメッセージ通知: 予約・キャンセル完了時に生徒のLINEトークへ通知を送信。

## 🛠 テックスタック
- Front-end: HTML5, Vanilla JavaScript, CSS3
- Back-end: Google Apps Script (GAS) / Node.js (clasp による開発)
- Database & Storage: 
  - Google Sheets: メインデータベース（永続データ管理）
  - Cloudflare Workers KV: レスポンス向上のための高速キーバリューストア（キャッシュ層）
- Calendar: Google Calendar API
- Platform: LINE Developers (LIFF, Messaging API)
- Deployment: Cloudflare Pages

## 📂 ディレクトリ構成  
├── gas/                 # バックエンド (clasp / Google Apps Script)  
│   ├── node_modules/    # 型定義等のライブラリ (Git除外)  
│   ├── .clasp.json      # clasp 接続設定 (Git除外)  
│   ├── appsscript.json  # GASプロジェクト設定  
│   ├── Code.js          # メインロジック  
│   ├── ExceptionForm.html # エラー表示用UI  
│   ├── migrate.js       # データ移行・スライド処理用  
│   ├── package.json     # Node.js依存管理  
│   ├── Properties.js    # 環境変数・プロパティ管理  
│   ├── SpreadSheetCustomMenu.js # スプレッドシート上のカスタムメニュー  
│   ├── syncDataToWorkers.js     # Cloudflare Workers KV 同期ロジック  
│   └── User.js          # ユーザー管理ロジック  
├── liff/                # フロントエンド (Cloudflare Pages)  
│   ├── index.html       # 予約システム画面  
│   ├── script.js        # LIFFロジック（ビルド時にURL置換）  
│   └── style.css        # スタイルシート  
├── .gitignore           # Git除外設定  
└── README.md            # プロジェクト説明書  

## 🚀 セットアップとデプロイ  
後で更新予定

## ⚙️ 定期メンテナンス (トリガー)  
後で更新予定

## 📊 シート構成  
後で更新予定

## 💡 開発者メモ
- 軽量化: 履歴データは保持せず、削除する方針です（スプレッドシートの行数制限対策）。
- UI/UX: 「今日」の日付には視認性を高めるためCSSで丸い背景を適用。
- 整合性: 月を跨ぐ予約の整合性を保つため、スライド処理は1日の早朝に完結させる必要があります。
- ローカル開発: clasp を使用し、VS Code上で開発。@types/google-apps-script による型補完を推奨。

## ⚖️ ライセンス
個人利用および特定のコミュニティ内での利用を想定しています。
