# A列・B列の別行番号への対応

2026-10-06の依頼に対応するpreview.39。A5＝1・B6＝1を6行目のCASE「1-1」として認識する。

## 変更

番号の共通正規化処理で、A列の有効な数値をB列が空でも保持する。B列の有効な数値が現れた行をCASE開始行にする。後続のA列番号で大番号を更新し、同じ行のA・B番号も従来どおり使用する。A列番号が一度もないB列番号はCASEとして採用しない。非数値の説明行は大番号を更新せず、重複CASEは従来どおり拒否する。

最初のB列番号の直前にSide見出しがない場合は、その大番号を指定したA列行の直前を読む。これによりNEW／OLD見出しが4行目、A5＝1、B6＝1の配置にも対応する。CASE範囲・自動配置・CASE移動は共通の認識結果を使い、セルの番号を書き換えない。

## 検証

Releaseビルドは警告0・エラー0。非Excelテスト172件が成功。A列のみの開始・大番号変更・全角数値・空欄の引継ぎ・非数値・従来形式・正規化の再適用を確認。

影響する実Excelの4シナリオが成功、スキップ0。既存のSnapshot読取・同Side後追い・画面自動配置／Undo／Redo／次CASE移動の3件は `artifacts/split-case-numbering/excel.trx` に記録。追加した `SplitCaseNumbering_ResolvesPlacesNavigatesAndReplays` は初回に一時ブックが探索一覧に現れず接続確立前に失敗し、単独再実行で成功した（`split-excel-final.trx`）。

追加試験ではCASE「1-1」「1-2」「2-1」「3-1」の開始行6・15・26・40、1-1の終端14行目、NEW／OLD見出し、自動配置、実座標を用いた削除・再配置、前後CASE移動を確認した。A列行の直前・B列行の直前の両方の見出し位置で解析でき、A5の値が変わらないこと、重複CASEを拒否することも確認した。専用Excelプロセスと一時ブックを使用し、報告元ブックは操作していない。

## 発行

改修ソース: `bfe889405b9dbf55485ed675a6ffdbb1f55156bb`。
`bootstrap.ps1 -Publish -Runtime win-x64` が成功。最終Releaseビルドは警告0・エラー0、非Excelテスト172／172成功。単一EXEとSHA-256 sidecarの一致を確認した。

- EXE: `artifacts/publish/win-x64/EvidenceCrafter.exe`
- ProductVersion: `0.1.0-preview.39+bfe889405b9dbf55485ed675a6ffdbb1f55156bb`
- SHA-256: `259C9F78A79FAFCA6ADF165222A22599A776C825A11DC1F28EBCC37A34C90613`
- 発行ログ: `artifacts/split-case-numbering/publish.log`

preview.38のEXEは `artifacts/split-case-numbering/EvidenceCrafter-preview38.exe` に退避した。ソースと検証記録をremote mainへ反映し、EXEと生記録は既存方針どおりignored artifact。
