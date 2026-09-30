# 「同じSideの次CASE」新側先行・旧側後追い配置の改修計画

作成日: 2026-09-30。調査対象: `47acd10` 時点のコード。
ステータス: 改修実装済み（2026-10-01）。再現、実Excel回帰、性能測定と配布結果は [改修検証記録](review-same-side-backfill-2026-10-01.md) に記載する。以下は実装前に作成した計画であり、当時の調査結果と目標を残す。

目的は、NEWを複数CASEへ先に配置し、OLDを後から配置しても、対応画像の上端を揃え、画像数の増加に伴う待ち時間を減らすこと。
「横並び」は同じCASE・同じ段の画像の実際のTop座標が揃うこととする。開始セルの行番号が同じだけでは合格にしない。

## 1. 調査で確認できたこと

| 箇所 | 確認した実装 | 今回への影響 |
| --- | --- | --- |
| `MainForm.ShowImagePreviewAsync` → `PlaceClipboardImageAutomaticallyAsync` | プレビュー解析、配置、両側完成時の末尾整理、次CASE移動を順に実行する | OLD後追い時は参照画像処理と行整理が増える。貼り付け単体と操作全体を分けて測る必要がある |
| `ExcelAutomaticPlacementService.FindPair` / `CaseImages` | 同一CASE内の管理画像を開始行・Top・名前順に並べ、同じ順番の反対側画像を参照する。再解析時は参照画像名を引き継ぐ | 対応付けの仕組みは既存。単純にSide別の末尾へ置いているわけではない |
| `AnalyzeSnapshot` / `PlacementPlanner.Plan` | 対応画像の開始行を `PreferredStartRow` に渡す。移動が必要な場合だけ `TargetTopPoints` を計算する | セル内の縦オフセットを通常の貼り付け計画へ渡していない |
| `ExcelImagePlacementService.PlaceWorkbook` | `AddPicture` のTopは常に `cellTop + 2pt` | 参照画像がセル内の別の位置にあれば、新規画像と実上端が一致しない |
| `PlaceImages` | 通常のペア配置は、サイズ・位置が同じでも参照画像更新経路へ入り得る。その後 `PlaceImages` を再帰呼出しする | Inspect、Resize内の全Shape衝突確認、Analyze、Snapshot照合が重なる |
| `AnalysisMatchesRequest` | 計画内の画像レコードと入力画像レコードを `SequenceEqual` で比較する。一方、解析で `ReferenceShapeName` を付加している | ペアのある未編集プレビューでも入力レコードと不一致になり、再解析へ進む条件がある |
| `SnapshotStillMatches` | Snapshotを新規取得して真偽値のみ返す | 差異があった場合、直後のAnalyzeが再取得し、取得済み情報を使えない |
| `ExcelSheetSnapshotService.CaptureWorkbook` / `ReadShapes` | 占有セル・行高は対象CASE中心だが、Shapeは全シート分のセル座標・寸法・管理情報を毎回読む | 無関係なCASEの画像増加も配置解析のCOM通信量に影響する |
| `TrimCompletedCaseTail` / `ExcelRowMutationService` | 配置後に完全Snapshotを取得し、削除直前にも全Shapeを走査して行の安全性を確認する | 後処理も性能測定の対象。ただし削除直前の安全確認は必須 |
| `ExcelCaseNavigationService.Navigate` | 既に `CaptureForNavigation(... includeShapes: false)` を使用する | 次CASE移動のShape読み取り省略は対応済み |

位置ずれの有力な原因候補は「実Top」と「行＋固定2pt」の不一致である。
複数CASEにまたがる行挿入・末尾削除後のExcel実座標変化、ペア選択の変化、衝突時の退避も調べ、累積的なずれの直接原因は再現結果で確定する。

既存の `review-capture-editor-window.md` では60Shapeの読み取りが約559ms、配置解析が約779msだった記録がある。
これは過去の別条件の測定であり、現在の速度や今回のボトルネックを確定する値ではない。COM往復が有力候補であることを裏付ける参考値として扱う。

## 2. 配置の契約

- 対応付けは既存の「同じCASE内の同じ順番の管理画像」を使用する。処理中の再解析では確定した参照画像名を引き継ぐ。
- 両画像の上端差は、既存の実Excelテストに合わせて **0.05pt以下** とする。後続CASEでもこの許容差を超えないこと。
- 行挿入・削除後は最新の参照画像の実Topを基準にする。前CASEの座標や過去の行高から累積加算しない。
- 対応画像より大きい画像を置く場合も、段の高さは両画像の大きい方で確保する。後続画像との2行間隔とCASE境界を守る。
- 同じ段に既存セルや図形があれば、既存のペア移動・行追加を使って安全な共通位置を確保する。確保できなければ停止し、片側だけを下へ逃がして成功とはしない。
- 管理外画像や対応相手のない画像は、既存の占有判定・追記規則に従う。管理外画像を見た目でペアと推定しない。
- 共通表示幅・回復時の参照サイズ保持、NEWのみ、反対Sideへ進むモード、手動配置、Undo/Redoを回帰対象にする。

## 3. 実装順序

### 段階1: 再現条件と計測を用意する

既存の `PairedImageIntegrationChecks.cs`、`ReferenceAppendChecks.cs`、`PlacementReadPerformanceChecks.cs` を拡張する。新しいテスト基盤は追加しない。

1. 生成ブックで `NEW 1-1 → NEW 1-2 → … → NEW 1-N → OLD 1-1 → … → OLD 1-N` を実行する。実アプリと同様に各OLD配置後の末尾整理と同一Sideの次CASE移動も含める。
2. 各配置、参照サイズ変更、行挿入、末尾整理の前後でCASE境界、参照Shape名、実Top、開始・終了行を記録する。どの段階で差が生まれたかを切り分ける。
3. 均一行高に加え、小数行高、混在行高、非表示行、参照画像のセル内移動、縦長OLDによる行追加を組み合わせる。同一CASEに複数枚ある既存シナリオも残す。
4. `Stopwatch` と既存診断ログを使い、プレビュー解析、確定時再検証、参照画像処理、行追加と再解析、画像追加、末尾整理、次CASE移動を別々に測る。
5. Snapshot取得回数、全Shape列挙回数、Shape詳細取得件数、参照画像更新回数、再解析回数も記録する。低レベル通信回数が必要なら対象のCOM呼出ラッパーで限定的に計測する。

成果物: 最小再現ケース、処理別時間・呼出回数の基準値、失敗する位置合わせの回帰チェック。
現象そのものを再現できない場合は、実Topが固定2ptから外れるケースを欠陥の再現として扱い、ユーザー報告の累積ずれとは区別して記録する。

### 段階2: 計画・配置・履歴で同じ実座標を使う

主な変更先: `PlacementModels.cs`、`SheetSnapshot.cs`、`PlacementPlanner.cs`、`PairedImageResize.cs`、`ExcelAutomaticPlacementService.cs`、`ExcelImagePlacementService.cs`、`MainForm.cs`。

1. 既存のペア計画へ参照画像の実Topを保持し、貼り付け側まで明示的に伝える。行内オフセットが必要な場合はSnapshot取得時のセルTopとの差を保持し、CoreでExcelを読まない。
2. 必要行数は固定2ptではなく、実際に使う行内オフセットと段の高さから計算する。通常の単独配置の既定値は2ptとする。入力は有限値・有効範囲を検証する。
3. 参照画像を移動する場合は、参照画像と追加画像へ同じ目標Topを渡す。行追加後は実座標を読み直して最終計画を作り直す。行高15ptの仮定だけで最終座標を確定しない。
4. 貼り付け結果の実Top・寸法・CASE内収容・非重複を検証する。ペア上端差が許容差を超えた場合は、既存の補償経路で新規画像・挿入行・参照変更を戻す。
5. `AddAutomaticPlacementHistory` 内の再配置も、確定済みのTopまたは行内オフセットを使う。Redoで再び「行＋2pt」に戻らないよう、履歴内の復元順序に合う座標を保持する。
6. 実装時に `PlaceImage` の全呼出元を更新・確認する。自動配置とそのRedoへ指定値を渡し、手動配置の既定動作も検証する。

成果物: 段階1の位置合わせ回帰チェック成功、通常・回復経路・Undo/Redoで同じ配置結果。

### 段階3: 重複する更新とSnapshot取得を削る

主な変更先: `ExcelAutomaticPlacementService.cs`、`PairedImageResize.cs`。必要な範囲だけ `ExcelManagedShapeService.cs`。

1. 入力の一致判定から解析が内部付加した参照画像名などを分離する。比較すべき利用者入力と解析結果を区別し、画像・寸法・Sheet・CASE・Side・余白など、計画に影響する条件の不一致は確実に検出する。
2. 確定時に最新Snapshotを一度取得し、照合に使用したSnapshotをそのまま `AnalyzeSnapshot` に渡せる構造へ変える。不一致だからという理由だけで同じ状態を直後に再取得しない。
3. 参照画像の実位置・寸法・必要メタデータ・追加行に変更がなければ、Resizeと変更履歴作成、変更後の再帰解析を省く。最新状態と衝突の確認は配置側で行う。
4. 実際に参照変更や行変更がある場合だけ、その変更後の再読込を行う。既存の回復・補償・リトライ上限を維持し、変更前Snapshotを変更後の検証へ使わない。
5. プレビュー中のExcel手編集を検出する確定時の再検証は残す。処理間や複数操作をまたぐ永続キャッシュは導入しない。

成果物: 状態が変わらない通常ペアで参照Resizeが0回、Snapshot差異時の照合直後の二重取得が解消され、外部編集の回帰テストも成功。

### 段階4: 画像数に比例するCOM読み取りを軽くする

段階3の再計測後、残る全Shape読み取りを改善する。主な変更先: `ExcelSheetSnapshotService.cs`。

1. 配置用Snapshotでは全Shapeの最新境界を取得して対象CASEとの交差を判定し、詳細な管理情報・寸法などは配置判定に必要なShapeを中心に読む。両SideをまたぐShape、CASE外から入り込むShapeも含める。
2. 詳細情報を読まなかったShapeを、正常な0ptの画像として扱わない。既存APIの完全Snapshot契約と配置用の取得範囲を区別し、Fingerprintも同じ取得契約どうしで照合する。
3. CASE未指定時の空きCASE探索、CASE外へ移動した名前指定参照、Undo/Redoなど、全体情報や追加情報が必要な呼出元には完全取得または明示的な追加取得を用意する。
4. 削除直前の全行・全幅の安全確認は維持する。行削除はCASEの列外にも影響するため、配置用の限定Snapshotで代用しない。
5. 必要なら、末尾整理前の重複取得やWorkbook再接続を、同じSTA処理内で寿命を限定して減らす。これは計測で支配的と分かった場合に限定し、汎用COMセッション基盤の全面改修には広げない。

全Shapeの現在位置を確認する走査は残るため、この段階でも総Shape数Nに対してO(N)である。狙いは詳細なCOM呼出数と同じ走査の繰返しを減らすこと。
VBA導入やShapes全プロパティの一括転送を前提にはしない。C#側の計算並列化も、COM往復削減の代替にはしない。

## 4. 検証と受入基準

| 観点 | 検証内容・合格条件 |
| --- | --- |
| 実操作順序 | 複数CASEのNEW先行→OLD後追いを末尾整理・次CASE移動込みで実行。各対応画像のTop差が0.05pt以下 |
| 累積ずれ | 各操作直後と全CASE完了後の両方で比較。先頭・中間・末尾CASEで許容差超過がない |
| 段・サイズ | 同一CASE複数枚、異なる縦横比、参照画像拡大、サイズ保持の回復で、上端一致・2行間隔・非重複を確認 |
| 行と既存内容 | 小数・混在・非表示行、行追加と自動整理、管理外Shape、コメント・数式・リンクがある場合も既存内容とCASE境界を保護 |
| 状態変化 | プレビュー後の画像移動・削除・改名・追加、行高・列幅変更、行挿入、ブック再オープン、保護を検出し再解析または安全停止 |
| 履歴・失敗 | Undo/Redoで位置・サイズ・行を復元。失敗途中の補償と、利用者が編集した行の削除拒否を確認 |
| 他モード | OLD先行、同じCASEの反対Side、NEWのみ、管理外参照、ペアなし、手動配置の回帰なし |
| 参照ブック | 一時コピーだけで検証し、参照元のSHA-256が不変 |

性能は同一PC・同一Excel・同一生成条件で変更前後を比較する。
既存画像数0／50／200／500を基本とし、対象CASE内の画像数を固定して他CASEだけ増やす条件と、同一CASE内を増やす条件を分ける。
通常ペア、参照拡大、行追加・末尾整理ありを別々に測る。ワークシートのサイズ上限も既存Snapshot制限内に収める。

- プレビュー表示まで、配置確定から次CASE選択完了まで、工程別時間を別々に報告する。利用者の確認待ち時間は含めない。
- ウォームアップ後、同じ初期状態のコピーを使って10回以上測り、中央値と最大値、Shape列挙・Snapshot回数を残す。連続したOLD後追い全体の時間も測る。
- 初期の性能目標は200／500画像条件の通常ペアで確定後の総時間中央値を30%以上短縮、小規模条件で10%超の継続的悪化がないこと。これは実測前の目標であり改善率の保証ではない。
- 達成しなければ工程別の残存時間から追加変更を決める。画像読み取りだけの改善率を、操作全体の改善率として報告しない。

実装時の基本確認コマンド:

```powershell
dotnet build EvidenceCrafter.sln -c Release
dotnet test tests/EvidenceCrafter.Tests/EvidenceCrafter.Tests.csproj -c Release --filter 'TestCategory!=ExcelIntegration'
dotnet test tests/EvidenceCrafter.Tests/EvidenceCrafter.Tests.csproj -c Release --filter 'TestCategory=ExcelIntegration'
```

実ExcelテストはOffice導入環境で生成ブックと参照コピーを使用する。通常テスト全件成功、Release build警告0、実Excel受入成功、変更前後の計測結果を記録して完了とする。
配布物を更新する実装タスクでは、既存手順に従って `bootstrap.ps1 -Publish` と参照ベースライン確認も行う。

## 5. 実装単位と完了時の報告

1. 再現・計測と実Topによる位置合わせを一つ目の変更として完成させる。
2. 不要な参照更新・重複Snapshot取得の削減を二つ目の変更として計測する。
3. Shape詳細取得の範囲縮小を三つ目として検証する。さらに必要な最適化は測定結果に基づいて限定する。

各単位で、実Top差、対象CASEの保持、既存内容の保全、Undo/Redoを確認してから次へ進む。
完了報告には再現できた原因、位置合わせ結果、画像数別の操作時間、削減したCOM処理、未実施の実機確認があればその範囲を載せる。
今回の計画作成では製品ソースコードの変更とExcelブックの操作は行っていない。Release構成でExcel非依存テストを実行し、161件が成功、失敗・スキップは0件だった。結果は `artifacts/review-results/same-side-plan-baseline.trx` に保存した。今回の報告事象の実Excel再現、画像数別の性能測定、改修後の検証は未実施であり、段階1以降で実施する。

## 6. 実装時に守る境界と補足

- 参照専用の `C:\work\Macro\Case&Evidence` と利用者ブックは保存しない。実装対象は存在を確認した `C:\work\EvidenceCrafter`。環境情報の `C:\work\CraftEvidence` は現在存在しない。
- 行移動後のメタデータのAnchorCellを実座標の代わりに使わない。TopLeftCellとShape.Topの現在値を使用する。
- 行内オフセットは `Shape.Top - TopLeftCell.Top` とし、非表示行・小数行高を含めて同じオフセットで計画と実行を一致させる。行追加・参照移動が終わった時点の実Topを最終基準にする。
- 初期の位置合わせチェックは全CASE完了後にも再実行する。後から前のCASEの行が削除された際、既に配置済みの両画像が異なる量だけ動く問題も検出する。
- 画像のないCASE、複数画像のCASE、途中からOLDを追加する操作を含め、CASE番号と確定した参照Shape名で追跡する。CASEを一定行数や前CASEの累積高さで識別しない。
- 配置用の限定Snapshotを追加する場合は、取得範囲と情報の完全性をDTOに明示する。別CASEから範囲内へ移動したShape、範囲外からはみ出すShape、両SideをまたぐShapeを最新状態で拾い、境界上で判断できない図形は詳細取得へ進む。元の完全Snapshot APIの意味を暗黙に変更しない。
- 既存のValue2／Formulaの範囲一括取得、行高の均一ブロック取得、Shapeセル番地の一括解析、移動時のShape省略は既に実装されている。それらを今回の新規改善として数えない。
- Focusの呼出回数も測る。現在は画像追加サービスのFocusと、その後の次CASE移動のFocusが別々に実行される。次CASE自動移動の経路で重複が支配的なら、追加サービスの既定動作を維持したまま内部指定でFocusを後段へ委譲し、移動失敗時には配置セルへ戻す。プレビューの配置先表示と手動配置のFocusは別に検証する。
- `EnableEvents` と `ScreenUpdating` は変更前の値を保持し、失敗を含めてfinallyで復元する。計算モードの変更は初期改修に含めない。Officeへの呼出しはSTAで直列に実行し、別ApartmentへRCWを渡さない。操作単位の参照共有を導入する場合も、ユーザーのプレビュー待ちをまたいで保持しない。

技術判断の参考:

- Shape.Topはシート上端からのポイント座標であり、実上端の比較基準にできる。[Microsoft: Shape.Top](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.top)
- MicrosoftはExcelとのデータ転送回数を減らすこと、セル値は範囲単位で読み書きすること、変更したExcel設定を元の状態へ戻すことを推奨している。今回も必要な再検証を維持して重複転送を減らす。[Microsoft: Excel performance](https://learn.microsoft.com/en-us/office/vba/excel/concepts/excel-performance/excel-tips-for-optimizing-performance-obstructions)
- Officeのオブジェクトモデルはスレッドセーフではなく、COM呼出しはSTAで直列化される。今回の高速化は主に呼出回数と取得情報量の削減で進める。[Microsoft: Threading support in Office](https://learn.microsoft.com/en-us/visualstudio/vsto/threading-support-in-office?view=vs-2022)

Routing: COMPLEX -> strong capability
