# 同一CASEへの連続配置と負のオフセット（2026-11-10）

## 再現

専用Excelプロセスと一時Workbookで、同一CASEにNEW画像を6枚配置してからOLDを後追いする条件を検証した。通常の行高13.5ptでは、同じ大きさ／OLDだけ縦長の各6枚は変更前も成功した。

負値の自然発生を確実に捉えるため、別のfixtureでは行高409.5pt、200000行付近に管理画像を6枚用意し、その前方に5行ずつ4回挿入して追加行を15ptへ変更した。Excelが返した `Shape.Top - TopLeftCell.Top` は −2.5、−0.5、9.5、405、−2.5、7.5pt。値を注入せずに、最初のOLD配置で報告と同じ `The vertical offset must be finite and no less than -1pt. (Parameter 'offset')` が発生した。

セル座標と画像座標の精度の違いにより、行が移動すると表示上の上端とアンカーセルの上端に負の差が生じる。従来はこの差をそのまま配置計画の入力として検証していた。報告元のWorkbookは未入手であり、この大きな座標のfixtureが利用者の自然発生条件と同じだと確認したものではない。

## 修正

後追い配置は既存画像の実際のTopを使う。行数の計画では負のオフセットを0として必要領域を保守的に確保し、画像の配置自体には補正前の実Topを渡す。セル上端へ丸め直した位置を次の基準にせず、参照画像を不要に移動させない。行追加後は従来の再解析で更新されたTopを使用する。

自動配置のUndo／Redoも、対応する参照画像の状態確認と行復元を行ったうえで履歴の実Topを使用する。セル境界付近では挿入画像の実TopLeftCellが予定セルと異なるため、履歴には実TopLeftCellも取得して保存する。これにより、配置直後なのに外部編集と誤判定して削除・Undoを拒否することを防ぐ。非有限値や負の絶対座標は拒否する。単独配置のオフセット入力の制限、位置・サイズの1pt許容、外部編集検出、衝突回避・行の安全確認・補償は維持する。

## 回帰テスト

- `PairedImages_BackfillAfterRowInsertionsWithNegativeOffsets`: 実Excelで負値を自然発生させ、同一CASEの6組を配置。参照画像の上端不変・余分な行追加がないこと・最初の組の削除と再配置を3回繰り返してもずれが増えないことを確認。
- `PairedImages_BackfillSixImagesWithoutDrift`: 同一CASEにNEW6枚→OLD6枚。OLDの同サイズ、縦長、幅・高さとも大きい3条件で、各配置後の全組の上端、行追加、CASE末尾整理を確認。同サイズでは既存画像が動かないことも確認。
- `AnalyzeSnapshot_UsesActualPairTopAndReservesSpaceForNegativeOffsets`: 参照サイズ保持の有無それぞれで負値・0・正値の計画と実Top保持、NaN／Infinityの拒否を確認。
- `AbsoluteTop_RejectsInvalidCoordinatesBeforeAccessingExcel`: 絶対座標の入力検証。

一括実行で、既存の画面操作テストがMainFormをDisposeしたあと接続レジストリの終了状態を残し、後続のExcel探索・レジストリ単体テストが失敗することも分かった。画面操作fixtureの後始末でテスト開始前の終了状態を復元し、実行順への依存を除いた。製品の終了処理は変更していない。

最終コードで通常テスト180件、実Excel9件がすべて成功、スキップ0。既存のペア拡縮・衝突回避・複数CASEの後追い・丸め許容と補償・画面操作経由のUndo／Redoも含む。参照資料6件のSHA-256はベースラインと一致。

記録は `artifacts/backfill-2026-11-10/` 配下（Git対象外）。

- `negative-before.trx`: 変更前に報告と同じエラーを再現。
- `final-unit.trx`: 通常180件。
- `final-sequence.trx`: 新規4条件と画面操作テストの計5件。
- `final-placement.trx`: 既存の実Excel配置4件。
