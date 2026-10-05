# 配置の0.1pt丸めと1pt許容

ユーザー依頼「大体一緒ならOK」に対応するpreview.38。

## 変更

- `PlacementGeometryComparison` に比較を共通化。双方を0.1pt単位へ四捨五入し、その差が1pt以内なら受け入れる。中間値はゼロから遠ざかる方向に丸める。たとえば77.04ptと78.049ptは77.0ptと78.0ptになり成功、78.051ptとの比較は差1.1ptで拒否する。
- 新規画像のLeft／Top／Width／Height、NEW／OLDの上端合わせ、参照画像の倍率変更後の検証、配置可能幅、倍率変更時の図形重なりと行高さ不足の判定に適用。
- サイズの再設定による補正とSingle隣接値の特別扱いを削除。初回の実geometryが許容範囲内ならその値で成功とする。
- 履歴には丸める前のExcel実geometryを保存。履歴対象がユーザーに編集されていないかの比較は従来の0.05ptを維持する。名前・接続・セル内容・CASE境界・取り消し失敗の保護も維持。
- 失敗診断に比較精度0.1pt、許容差1pt、丸めた予定／実座標を記録する。

## 検証

Releaseビルドは警告0・エラー0。非Excelテスト171件が成功した。丸め境界、有限値・正サイズ、幅上限、診断の実値保持を確認。

影響する実Excelの7シナリオが成功、スキップ0。既存の6シナリオは `artifacts/rounded-placement-1pt/excel.trx`、追加テストの最終結果は `rounded-excel-final.trx` に記録した。追加テストの初回期待値は要求した座標差を用いていたが、Excelは値の設定時にさらに丸めるため、独立COM読取で得た実値から期待値を計算するよう修正した。製品の許容差は変更していない。

- 低い／100,000行目の座標で、Left／Top／Width／Heightの6種類の差を48条件確認。実値を0.1ptへ丸めた差で成功・拒否が決まること、履歴の4辺が独立COM読取の実値と一致することを確認。
- 報告値相当の幅・高さ−0.3ptを受け入れ、受入後の外部編集を従来どおり検出。画像サイズの差が許容内でも配置可能幅の超過が1ptを超える場合は拒否。
- NEW／OLDの実Top差1ptの受入、自動配置・後追い・Redo、参照画像の拡縮、同Side後追い、画像差し替え、行操作、画面Undo／Redo、次CASE移動を確認。
- 名前変更、大きな座標・サイズずれ、取り消し失敗、複数画像の途中失敗による画像・行の補償、既存画像・セル内容・CASEアンカーの保全を確認。

専用Excelプロセスと一時ブックで試験し、報告元ブックは操作していない。

## 発行

改修ソース: `0fb859656a4091cdf3fbf50c10585a971d3abbdf`。
`bootstrap.ps1 -Publish -Runtime win-x64` が成功。最終ビルドも警告0・エラー0、非Excelテスト171／171成功。単一EXEとSHA-256 sidecarの一致を確認した。

- EXE: `artifacts/publish/win-x64/EvidenceCrafter.exe`
- ProductVersion: `0.1.0-preview.38+0fb859656a4091cdf3fbf50c10585a971d3abbdf`
- SHA-256: `A56EF2DB402FCC211F1601CB38638235D351746D9A6FF4AFBB4FFEDB2DAFCE41`
- 発行ログ: `artifacts/rounded-placement-1pt/publish.log`

前版EXEは `artifacts/rounded-placement-1pt/EvidenceCrafter-preview37.exe` に退避済み（SHA-256: `5A02C63D46134B677C63B4B229B1F05E966E6549BF563C9BFC834E8D916C43F8`）。ソースと検証記録をremote mainへ反映し、EXEと生記録は既存方針どおりignored artifact。
