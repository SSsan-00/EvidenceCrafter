# Excel通信・末尾整理の追加高速化計画

作成日: 2026-10-01。対象: preview.31の現在の作業ツリー。
状態: 段階0〜5をpreview.32に実装し、単体・実Excel・比較計測・独立レビュー・publishを完了。200／500図形での総時間中央値は20.9%／26.6%短縮し、暫定30%目標は未達。安全確認を維持し、計画6の許容条件に従って改善分を採用した。[実装・検証記録](review-excel-com-performance-2026-10-01.md)に結果と残存時間を記録した。

## 1. 結論と狙い

追加高速化の余地はある。最初は、全Shape走査を維持しながら不要な終端セル取得を省き、Undo用fingerprintの取得を範囲単位にまとめる。安全判定の省略や永続キャッシュは使わない。改善後に再計測し、必要な場合だけ末尾整理の事前Snapshotと削除処理を統合する。

ただし、行削除後の履歴情報取得に失敗すると退避ファイルまで削除する既存経路があるため、その異常時の扱いを先に直す。高速化のために新しい読取経路を増やす前提条件とする。

## 2. 現在の根拠

前回の[改修検証記録](review-same-side-backfill-2026-10-01.md)から、preview.31の測定値を基準にする。

| 対象CASE外の図形数 | OLD配置の中央値・10回 | 末尾整理・1回 | 確定→末尾整理→次CASE移動・1回 |
| --- | ---: | ---: | ---: |
| 0 | 0.436秒 | 3.349秒 | 4.072秒 |
| 50 | 0.673秒 | 3.766秒 | 4.719秒 |
| 200 | 1.430秒 | 5.245秒 | 6.988秒 |
| 500 | 3.187秒 | 10.090秒 | 13.791秒 |

末尾整理には事前Snapshot、削除安全走査、退避ブック作成・保存、行削除、fingerprint、数式・名前定義・印刷範囲の取得が含まれる。前回はこれらを分離していないため、約10秒のすべてをバックアップやfingerprintの時間とは断定できない。

コードで確認した改善候補は次のとおり。

| 箇所 | 現状 | 改善候補 |
| --- | --- | --- |
| `ExcelSheetSnapshotService.ReadShapesCore` | 全ShapeのTopLeftCellとBottomRightCellを取得してから対象CASEとの交差を判定 | 開始行が対象より下なら、終端セルの取得を省く |
| `ExcelRowMutationService.ReadRowsWithShapes` | 削除安全確認でも同じ順序で両端セルを取得 | 同じ早期除外を適用。全列・管理外図形の保護は残す |
| `CaptureRowFingerprint` | 対象セルごとにValue2、Formula、表示形式、配置、Font、Interior、Validationなどを読む | 値・数式の一括読取と、均一と証明できる書式の共有 |
| `ExcelCaseMaintenanceService.TrimCaseTail` | 両Side確認用のSnapshotでShape走査後、別呼出しの行削除で再走査 | 同じWorksheet上で構造確認と削除安全走査をまとめる余地 |
| `CaptureNativeRowSnapshot` | `Workbooks.Add()`で退避ブックを作り、全行をCopyしてSaveAs | シート1枚を明示して、利用者設定による不要シート作成を防ぐ |

`CaptureFormulaMap`と行内容のValue2／Formulaは既に範囲一括取得済み。これを新規改善として数えない。SHA-256計算より、Excelプロパティの取得回数を減らす方を優先する。

## 3. 実装順序

### 段階0: 削除後の例外で退避を失わないようにする

対象: `ExcelRowMutationService.MutateWorkbook`、`RowMutationResult`、`RowDeletionSnapshot`、末尾整理結果を受け取る`MainForm`の呼出元。

現在は`targetRows.Delete`後の`PostDeleteFingerprint`等で例外が起きても、catchが`deletionSnapshot.Dispose()`し、`Changed=false`の失敗結果を返す。削除済みの行と、削除前の退避ファイルを区別して扱えていない。

1. 行変更前・変更完了・変更結果不明を区別する。DeleteのCOM呼出し自体が失敗した場合も、未変更と決めつけない。
2. 削除前の失敗なら通常の退避回収を行う。削除後または変更結果不明なら、復旧が確認できるまで退避ファイルを保持する。
3. 削除後の検証が不完全なSnapshotを通常のUndo/Redo可能な履歴として登録しない。復旧に必要なパスと実行結果を返し、末尾整理の失敗を表示して自動次CASE移動を止める。
4. 退避の保持だけでなく所有者・寿命も扱う。履歴20件上限、アプリ終了、通常のDisposeで未解決の復旧用退避を消さない。正常に復旧できたものだけ通常回収へ戻す。
5. 自動配置、画像差し替え後、画像削除後、手動行削除、Undo/Redoの結果処理を確認する。元の画像操作が済んでいる場合は、画像まで未変更だったという結果を返さない。

削除直後・fingerprint取得時・数式取得時の失敗を注入し、退避が残ること、未確認の復元を実行しないこと、復旧不能時に続行しないことを検証する。バックアップ作成・保存失敗ではDeleteを実行しない。

### 段階1: 末尾整理を工程別に測る

対象: 既存の`SameSidePerformanceChecks`と行変更サービス内の限定した計測箇所。

- 事前Snapshot、Shape走査、値／数式・コメント・リンク・結合確認、他Sheet数式確認を分ける。
- 退避ブックAdd、行Copy、SaveAs、Close、実際のDeleteを分ける。
- fingerprintと、削除前後の数式・名前定義・印刷範囲取得を分ける。
- プレビュー、画像配置、末尾整理、次CASE移動、確定から移動完了までの時間も残す。
- Snapshot回数、Shape開始／終端セル取得数、fingerprintのセル別取得数・範囲取得数・fallback数を記録する。論理的なCOM呼出数とRPC通信の実数を混同しない。

既存の内部observer／Stopwatchの形を使い、通常動作では計測を無効にする。新しい診断サービスや設定画面は追加しない。

比較前は現在のpreview.31ソース・ビルドを退避し、ハッシュを残す。`git archive HEAD`だけでは前のpreview.30相当になるため、今回の未コミット差分を含む基準版を固定する。

### 段階2: Shapeの不要な終端取得を省く

対象: `ReadShapesCore`と`ReadRowsWithShapes`。

1. ShapeのTopLeftCellと開始行を先に読む。
2. `start.Row > lastRow`なら、BottomRightCellを読まず対象外とする。
3. 名前指定参照がある場合は名前を確認し、一致したShapeには従来どおり終端・詳細取得を行う。
4. 開始行が対象CASEより上にあるShapeは除外しない。下側へはみ出す図形をBottomRightCellで確認する。
5. 対象内・境界上・名前指定参照の取得と、取得不能時の安全停止を維持する。

下にある無関係なShapeについて、終端セルオブジェクトとその番地／行の取得を省ける。今回の「最初のCASEをOLDで後追いする」条件に適している。後方CASEでは上にある図形が増えるため、同じ短縮率を期待しない。

全N個の最新状態の走査は残り、計算量はO(N)。狙いは1図形あたりのCOM通信量の削減。現在の`ReadShapeCellReference`を再利用し、座標キャッシュや新しい空間索引は作らない。

### 段階3: fingerprintの同値性を保って一括取得する

対象: `CaptureRowFingerprint`、`AppendFingerprintProperty`、既存`MatrixValue`。

1. 現在と同じ対象行・UsedRangeの使用列範囲について、Value2とFormulaをそれぞれ一度読む。25万セル制限は維持する。
2. `MatrixValue`を使い、現在と同じ行→列→プロパティ順序、`InvariantCulture`、区切り、SHA-256を維持する。単一セル・配列の下限・空欄・エラー値を正しく扱う。
3. NumberFormatは範囲から有効な文字列、WrapTextはboolが取得できた場合だけ均一値として共有する。混在、null／DBNull、取得失敗は該当プロパティをセル別読取へ戻す。
4. 初期実装では範囲一括＋セル別fallbackまでとする。混在した大きな範囲の分割は、計測で必要になった場合だけ既存`ReadRowHeightBlock`の考え方を使う。
5. Style、Font、Interior、Validation、FormatConditions.Countを「範囲の値が返ったから均一」と扱わない。初期実装では既存の個別読取を残す。NumberFormatとWrapTextは均一／混在の返値が公式に定義されている。[NumberFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.range.numberformat)、[WrapText](https://learn.microsoft.com/en-us/office/vba/api/excel.range.wraptext)。
6. 一括読取で失敗したときは既存の個別読取へ戻す。失敗を0／false／空欄などの有効値へ置き換えない。既存の非対応プロパティ表現との同値性も確認する。

既存のセル別実装をテストの比較基準として残し、同じ静的fixtureでhashの完全一致を確認する。fingerprintが検出している既存プロパティ集合を削らない。現在のfingerprintは全書式・全ルールの完全な比較ではなく、例えば条件付き書式はCountのみなので、その範囲を超える編集検出まで保証したとは報告しない。

値・数式・NumberFormat・WrapText以外の読取は残るため、改善率は実測で判断する。

### 段階4: 退避ブックを必要な1枚に限定する

対象: `CaptureNativeRowSnapshot`。

`Workbooks.Add`に`xlWBATWorksheet`を指定して、退避ブックをワークシート1枚に限定する。引数省略では利用者の`SheetsInNewWorkbook`設定に依存する。利用者設定は変更しない。[Microsoft: Workbooks.Add](https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.add)。

全行Copy、SaveAs完了後のDelete、ファイルによるUndo、数式・名前定義・印刷範囲の復元を維持する。既に新規シート数が1の環境では大きな速度差は見込まない。設定が1枚／複数枚の両方で、退避が1枚、復元結果が同じであることを確認する。

ここまで再計測し、十分な効果があればそこで完了する。

### 段階5: 必要な場合だけ末尾整理の重複取得を統合する

条件: 段階1〜4後も、末尾整理の事前Snapshot・同じ図形の重複走査が総時間の主要部分を占める場合。

対象: `ExcelCaseMaintenanceService`、`ExcelRowMutationService`、構造読取部分の`ExcelSheetSnapshotService`。

1. 末尾整理専用の要求に、CASE識別、tailRows、requireBothSidesを持たせ、既存の行変更実行経路へ渡す。
2. `MutateWorkbook`で対象Worksheetを解決した後、同じWorksheetからShape抜きの構造情報を取得し、既存`CaseLayoutAnalyzer`でCASEを解析する。構造読取は既存コードから必要な部分だけ再利用し、汎用COMセッション基盤を作らない。
3. 最新Shapeの安全走査で、全列の削除禁止行と、NEW／OLDの管理画像存在確認を同時に取得する。管理名・AlternativeText・実セル位置の判定条件は既存`HasManagedImage`と一致させる。
4. この結果で末尾削除を計画し、既存の全列安全確認・退避・履歴処理へ接続する。両Side確認用の事前Snapshot内のShape走査を省く。
5. 作成時メタデータのAnchorCellを現在の行位置として扱わない。実TopLeftCellと一意なCASE識別を使い、範囲や接続の変化を検出した場合は再解析または停止する。
6. 同じSTA・イベント抑制だけでは、Excel全体の排他や状態不変を保証できない。native backupなどの長い処理をまたいで古い安全判定を使わず、削除直前のライブ確認を維持する。追加の再確認が必要な経路は走査を残し、「全経路で1回」にすること自体を目標にしない。

最終CASE、NEWのみ、片Side未完成では従来どおり削除しない。自動配置後の2行維持と、手動・差し替え・画像削除後の4行維持を区別する。速度より同じ削除範囲・復元結果を優先する。

## 4. 初期改修に含めない案

- Shapeの永続キャッシュ: Countやセル変更イベントだけでは図形の移動・拡縮を検出できるとは限らず、削除安全判定を古い位置で行うリスクがある。
- VBA／XLM／追加アドインによるExcel側一括転送: 配布・実行条件が増えるため、今回の小さなCOM削減で不足することが実測された場合に再検討する。
- 退避のSaveAs省略、非同期保存、Undo用fingerprintの省略、JSONだけの書式復元: 現在の復元能力・失敗時の保存済み退避を損なう。
- Excel呼出しの並列化: Office呼出しはSTAで直列化される。RCWを別Apartmentへ渡して速度を得る方針は採用しない。[Microsoft: Threading support in Office](https://learn.microsoft.com/en-us/visualstudio/vsto/threading-support-in-office?view=vs-2022)。

## 5. 検証と受入基準

| 観点 | 確認内容 |
| --- | --- |
| 位置合わせ | 既存4CASEのNEW先行→OLD後追い、行整理・移動・Redo後の全ペアTop差0.05pt以内 |
| Shapeの同値性 | 対象より上／下、はみ出し、境界上、両Sideまたぎ、名前指定範囲外参照、管理外・回転・グループ図形、非表示行、小数行高で旧経路と同じ取得・削除禁止行 |
| fingerprintの同値性 | 均一／混在書式、NumberFormat、WrapText、文字・数値・0・false・空欄・数式・エラー値、単一セル、使用列の開始がA以外でセル別hashと完全一致 |
| 履歴の保護 | 値・数式・既存fingerprint対象の書式・行高・名前定義・印刷範囲の変更でUndo/Redoが従来どおり拒否される |
| 復元 | 行高・非表示・罫線・塗り・Font・入力規則・条件付き書式・依存数式・名前定義・印刷範囲を既存native backupから復元 |
| 削除安全 | エビデンス列外の文字・空文字数式・0・false、コメント・リンク・管理外Shape・結合セル・読取失敗を削除不可にする |
| 異常時 | backupのAdd／Copy／SaveAs失敗では削除しない。Delete結果不明・削除後読取失敗では復旧退避を失わず、未確認の続行・履歴操作を止める |
| Excel状態 | EnableEvents、ScreenUpdating、DisplayAlerts等を元へ戻す。計算モード・利用者の新規シート数設定・利用者ブックの保存状態を勝手に変更しない |

実Excelで旧経路と新経路を比較し、単体テストだけで書式の均一性や復元結果を証明したとは扱わない。現在の164件、既存Excelシナリオ、上記の追加チェックを実装段階に応じて実行する。

## 6. 性能測定と完了条件

1. 図形数0／50／200／500、対象CASEは前方・中央・後方に分ける。対象内画像数を固定して、均一書式と混在書式、削除なし／少数行／多数行を分けて測る。
2. 末尾整理を含む全体もウォームアップ後10回測る。前回の「全体1回」を今回の中央値の基準にしない。
3. 反復native backupでExcel切断があったため、既存supervisorを利用し、計測サンプルごとに専用Excelプロセスと同じ初期状態のfixtureを作る。起動・fixture生成・終了時間を操作時間から分離し、ユーザーのExcelを操作・終了しない。
4. 連続4CASEの操作も別に測り、初回だけの高速化や長期劣化がないか確認する。タイムアウト・切断・後処理失敗を記録し、成功した速いサンプルだけで合格にしない。
5. 中央値・最大値、工程別時間、読取回数とfallback件数を残す。計測前後で対象ソース・測定コード・fixture条件を固定する。

暫定目標は、現在のpreview.31と比較して200／500図形条件の末尾整理込み総時間中央値を30%以上短縮し、小規模・混在書式で10%超の継続的悪化がないこと。達成できなければ工程別時間を根拠に段階5を採否判断する。バックアップ自体が支配的で安全を維持した改善が乏しい場合は、効果のあった段階まで採用して残存時間・未達目標を明示し、安全確認を削って目標に合わせない。

計画の完了条件は、安全性・取得結果の同値性、必要な回帰テスト、比較計測、レビュー記録が揃うこと。30%は実測前の目標であり、速度保証ではない。製品コードと公開APIの変更は上記の必要範囲に限定する。

検証コマンドは既存のRelease build、`TestCategory!=ExcelIntegration`、対象の`ExcelIntegration`フィルター、`bootstrap.ps1 -Publish`を使用する。

Routing: CRITICAL -> strongest sufficient capability | Capabilities: current_documentation | Workers: review_backfill=unknown / unknown
