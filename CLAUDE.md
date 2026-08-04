\# 地域インフラ共創1課 中期計画データダッシュボード - プロジェクト概要



\## 概要

残業・有休・売上KPIをGASバックエンドと連携して可視化するダッシュボード。

URL: https://kasuyakouta.github.io/kankyoka-dashboard/



\## 技術スタック

\- フロント: GitHub Pages上の単一HTMLファイル(1,883行)

\- バックエンド: Google Apps Script (GAS)

\- GAS URL: localStorage(キー'kankyoka\_gas\_url')で管理、デフォルト値はindex.html 676行目に記載

\- グラフ描画: Chart.js 4.4.1(CDN経由、積み上げグラフ・ドーナツチャート)



\## このアプリ特有の設計(他アプリと異なる点に注意)

\- \*\*管理者PINはSHA-256ハッシュ化(Web Crypto API使用)\*\*。デフォルトPIN「3150」のハッシュ値をソースコードに埋め込み(平文は書かない設計)。localStorage(キー'kankyoka\_pin\_hash')に変更後ハッシュを保存

\- データ保存はlocalStorage即時反映 + GASへは2秒デバウンスで遅延送信

\- KPIカードはwarn/danger/ok/neutralの4状態で色分け表示、トレンド(up-bad/up-good/dn-bad/dn-good)も表示

\- モバイル専用の「残業・有休KPI開閉タブ」がPC表示とは別に存在



\## 必須ルール(標準スタック)

\- iOS Safari互換性・PWA対応を優先の設計制約とする

\- 日時はローカル時刻で組み立てる(UTCは使わない)

\- フォントは IBM Plex Sans JP / Noto Sans JP



\## 変更時のお願い

\- 複雑な変更は実装前にオプションA/B形式で提案し、承認を得てから実装する

\- 回答は簡潔に、前置きは省略する

\- PIN関連の変更時は「平文をソースに書かない」設計方針を維持すること

