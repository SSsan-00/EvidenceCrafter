# Top丸めによる画像配置取り消しの修正

対象: [実装計画](implementation-plan-top-verification-2026-10-02.md)。2026-10-03実装完了。配布版preview.35。改修ソース: `304f455e9f30563369c28da1f934500c0eb561f9`。

## 原因と変更

予定TopをSingleへ丸めた最寄り値と、ExcelがAddPictureから返す実Topが隣接Singleになる場合、現行版は正常な画像を取り消していた。計画作成時、専用一時ブックの72条件中6条件で同じTopエラーを2回再現した。通常オフセット2ptでも、予定937483.25ptに対する実値937483.3125pt、予定1874983.25ptに対する実値1874983.125ptを確認した。画像の再設定では差を解消できなかった。

共有の `PositionMatches` は従来の0.05pt判定を保持し、最寄りSingleから0.25pt以内の直前／直後Singleに実値が完全一致する場合を追加で許可する。標準の `float.BitDecrement`／`float.BitIncrement` を使い、無条件に許容差を広げない。幅・高さ、画像同士の実Top差、外部編集の検知、CASE末尾整理の安全条件は変更していない。自動・手動・画像Redoは同じ共有検証を利用する。

実Topを成功結果に保存する処理と、取り消しが確認できない場合の操作停止・退避保持を維持した。新たなCOM読取、全Shape走査、画像再挿入、待機や再設定ループは追加していない。判定は固定個数の数値比較だけで行う。

既存の失敗診断に最寄りSingleの予定Left／Top、追加丸め許容上限、Top位置検証結果を補い、アプリのInformationalVersionを記録した。既存JSONLローテーションを再利用し、Workbookパス・セル内容・画像内容を追加していない。

## 検証

- 単体170件が合格。6再現値の許可、通常位置のずれ・2つを越える丸め位置・0.25pt上限超過の拒否、座標0、不正値、サイズ0と厳密なサイズ比較、追加診断項目を確認。
- 実Excelの挿入画像fixtureに6再現条件を追加し、独立したShape.Top読取が調査時の実値と一致すること、削除・再配置で同じ実Topになることを確認。遠方列のLeftも確認した。
- MainFormを非表示でメッセージ処理し、絶対Topが大きいCASEでNEW先行→OLD後追いを実行する。正常な末尾整理と行未変更での整理中止の両方で次CASEへ進み、NEW／OLDの実Top差0.05pt以内を維持することを確認。実際のUndo／Redoも実行する。
- 初回fixtureではA:AFの200,202行へ行高を設定したため、後続試験がUsedRange安全上限に達して失敗した。製品の上限は変更せず、必要な202行だけを設定するfixtureへ修正。失敗記録 `artifacts/top-verification-fix/top-verification.trx` は保持した。

- Releaseビルド: 警告0・エラー0。最終発行時の単体テストも170／170合格、スキップ0。
- 全機能試験の初回実行は12件中11件合格。画面のUndo試験は計測上13時間を超えた後にタイムアウトし、生成Excelプロセスの強制終了を伴って失敗した。失敗原因を断定せず `integration-final.trx` に記録を保持した。
- 失敗した画面試験を単独で再実行し、1分54秒で後片付けを含め合格（`ui-retry.trx`）。既存の11件と合わせ、12種類すべての機能シナリオの成功を確認した。一括実行が12／12合格した記録ではない。
- さらにfixtureの印刷範囲を実CASEだけへ限定し、CASE間隔を16行にして、丸め再現に不要な印刷・大量行Undoの負荷を減らした。大きな絶対Topは維持している。最終fixtureで画像挿入と画面処理を再検証し、2／2合格（`affected-final.trx`）。挿入画像19.6秒、画面37.4秒。画面の実Undo／Redo後にも横並びを確認した。
- 検証終了後のExcelプロセス残存0。

原記録は `artifacts/top-verification-fix`。性能の再計測は、成功経路のCOM往復を増やしていないため今回は実施しない。高速化や測定済みの性能維持を主張する変更ではない。

報告元の具体的なTop数値・ブックは未提供。今回再現した丸めエラーを修正した結果と、報告元ブックそのものの解消確認は区別する。

## 発行

`bootstrap.ps1 -Publish -Runtime win-x64` 成功。単一EXEとSHA-256 sidecarだけの出力を確認した。

- EXE: `artifacts/publish/win-x64/EvidenceCrafter.exe`
- ProductVersion: `0.1.0-preview.35+304f455e9f30563369c28da1f934500c0eb561f9`
- SHA-256: `49EB861D4AEAE8843D1FD6F5FA63DBD707A1B5F1AB4F0C79631FE32B32D53647`（sidecar一致）
- 旧preview.34: `artifacts/top-verification-fix/EvidenceCrafter-preview34.exe`。旧hash `40DA9AF124B764CBADEFC2931A11C131385C2C0437933425A8564BD5EDBFB3CC` を確認して保持した。
- ソース・計画・検証記録をremote mainへ反映する。EXEと生記録は既存方針どおりignored artifact。
