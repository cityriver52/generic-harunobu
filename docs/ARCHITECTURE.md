# アーキテクチャ

## モジュール
- **Code.gs**：GASから呼ばれるWebアプリ入口、メニュー、認証済みユーザーのSheetsアクセス
- **Config.gs**：設定・処理上限・保存先の名称
- **Shared.gs**：安全なURL解析、A1表記、数値文字列等の共通判定
- **SheetReader.gs**：シート単位の読み取りと障害隔離、画像・非表示行列等の収集
- **RuleEngine.gs**：基本10項目の純粋な診断ロジック
- **RuleHeuristics.gs**：追加13項目のレイアウト・構造・値に関する参考判定
- **ReportWriter.gs**：生成した診断結果シートのみを変更する保存処理
- **Index.html**：フォーム、表示、フィルター、出典・ライセンス

```text
URL → Code.gs → SheetReader.gs → RuleEngine.gs + RuleHeuristics.gs
                                              ↓
                                        指摘配列
                                              ↓
                                   ReportWriter.gs（権限があれば）
                                              ↓
                                         Web画面
```

## 重要な設計方針
- 元のデータシートは変更しない。生成物として確認できない同名のシートは上書きしない。
- 共有URL以外の外部URLを読み込まない。Webアプリはアクセスしているユーザーの権限で実行する。
- シートごとの読み取り失敗は未検査として記録し、他のシートを検査する。
- シート名などのユーザー制御文字列を数式として実行させない。
- AI機能は使用せず、公式Harunobuの採点や合否を再現しない。

## テスト
`npm test`。NodeのVMが `tests/helpers.cjs` の定義順にGASスクリプトを読み込みます。GASへの実際のデプロイ・OAuth・画面動作の結合テストは別途必要です。
