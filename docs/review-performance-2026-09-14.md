# 実装レビューと性能改善（2026-09-14）

## 対象と修正

未pushだったSheet／CASE選択、フォーカス・最大化、テーマ色・透過度の変更を含め、
CASE移動と配置解析の呼出経路、設定保存、画面更新をレビューした。

- CASE移動は画像配置の有無に関係なく隣接CASEへ進む仕様だが、従来は毎回全図形を読んでいた。
  移動用の取得では `includeShapes: false` を指定して、このExcel通信を省略した。
- Sheet／CASE／Sideの選択に、行高・列幅・コメント・リンク・占有セルを読む完全Snapshotを使っていた。
  同じレイアウト情報を返す移動用Snapshotへ変更した。シート選択肢の取得でも図形は不要なので省略した。
- CASE選択肢が変化していなくても、ComboBoxを全消去して再追加していた。
  項目が一致するときは選択値だけを更新する。
- テーマの各色を参照するたびに色混合と輝度計算を繰り返していた。
  配色は基調色・強度が変わるときに再計算する。
- 中間色でも白文字を選択し、カードと背景の明度差も考慮しないため文字が読みにくくなる問題を修正した。
  各コントロールの背景に応じた白／黒文字と補助文字色を選ぶ。
- スライダーがマウスキャプチャを失ってもドラッグ状態を維持する問題を修正した。

配置の完全Snapshot、図形の重なり検証、ブックの接続・同一性確認は引き続き実行する。
画像を置く処理へ不完全な移動用Snapshotは渡さない。

## ベンチマーク

Windows、.NET 9、Release、実Excelに生成した一時ブックを使用。
240行×3 CASE、60図形（管理名30個／通常名30個）、A1:AF722という同じ条件。
各処理はウォームアップ1回の後、5回計測の中央値。完全Snapshot比較のみ3回。
元のExcelや参照ファイルを計測用に変更しない。

| 情報取得処理 | 従来の経路 | 改善した経路 | 時間短縮 |
| --- | ---: | ---: | ---: |
| CASE移動用Snapshot | 884.67 ms | 106.70 ms | 87.9%（約8.3倍） |
| Sheet／CASE／Side手動選択用Snapshot | 773.93 ms | 106.70 ms | 86.2%（約7.3倍） |

同じ改善後バイナリで従来経路（図形を取得する既定API）と軽量経路を比較した。
コード変更前の別計測でもCASE移動用Snapshotは981.82 msだった。
CASE境界、ヘッダーなどのLayoutSignalsが従来経路と一致することも毎回照合している。

画像配置用Analyzeは変更前1,005.33 ms、変更後986.53 msであり、
この差は実行負荷の変動として扱う。配置全体の高速化率としては報告しない。
配置計画のfingerprintは両方とも次の値で一致した。

`5C4EBFD05C1252F1B65E83B6ABF6AD3277B1C5E266A8F6FF55B68BC8F689D18E`

表の数字はExcel情報取得部分の値で、ボタン押下から画面描画完了までの総時間ではない。
CASE数・図形数・Excelの応答・PC負荷で効果は変わる。
テーマキャッシュやComboBox更新の改善は上記数値に含めない。

生データ（Git管理外）:

- `artifacts/review-results/before.trx`
- `artifacts/review-results/after.trx`

再現コマンド:

```powershell
dotnet test tests/EvidenceCrafter.Tests -c Release --artifacts-path artifacts/review-check --filter 'FullyQualifiedName~PlacementAnalysis_WithReal' --logger 'trx;LogFileName=performance.trx' --results-directory artifacts/review-results
```

## 回帰検証

通常テストに加え、実Excelの生成ブックと参照ブックのコピーを使う連携テストで確認する。
配色テストは8色×21段階で背景・カード・ラベル・選択欄・ボタンのコントラスト比を照合する。
スライダーテストは40～100%の範囲とキャプチャ喪失後に値が変化しないことを確認する。
最終結果は `artifacts/review-results/review.trx` に記録する。

今回の実行では通常テスト138件が全件成功した。実Excel連携は145件中144件が成功し、
`AppendImages_InReferenceCopies_DoNotOverlap` のみ、参照ブック4冊の検証後に生成Excelを終了した際の
STA終了待ちが180秒でタイムアウトした（配置結果・非重複検証までは成功）。同テスト単独の再実行でも
同じ終了待ちタイムアウトを再現したため、Excelプロセスの外部状態に依存する環境上の制約として扱う。
