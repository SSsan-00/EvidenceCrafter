# 小さな負の縦オフセットへの対応

preview.40で「The vertical offset must be finite and non-negative」の入力判定を修正した。

## 変更

`PlacementPlanner.NormalizeVerticalOffset` で有限値かつ−1pt以上を受け入れ、負の場合は0ptにする。正の値は変更しない。計画作成と画像挿入の双方から同じ処理を呼ぶ。NEW／OLD後追い、自動配置、直接挿入、Redoに適用される。

既存画像から読んだ `Top - TopLeftCell.Top` の値が小さな負値でも、必要行数と配置座標は補正済みの0ptで計画する。読取済みSnapshotや既存画像そのものは書き換えない。履歴は最終Excel実座標を使用し、上端の差は従来の配置比較で確認する。−1pt未満、NaN、正負Infinityは拒否する。

## 検証

Releaseビルドは警告0・エラー0。非Excelテスト173件が成功。−1pt・−0.3pt・微小な負値の計画が0ptと同じ行数になること、後追い参照画像の負オフセットでも解析に成功すること、正値の保持、非有限値と−1.001ptの拒否を確認した。

実Excelテスト3件が成功、スキップ0（`artifacts/negative-vertical-offset/excel.trx`）。−1pt／−0.3ptを直接指定した挿入がセル上端への配置になり、実座標による削除が成功すること、−1.001ptが拒否されることを確認した。NEW／OLDの上端合わせ、行追加後の参照画像復元、セル内容の移動と行Undoも既存シナリオで確認した。

報告元ブックでの自然発生原因を確定したものではなく、専用Excelプロセスと一時ブックでコード上の入力条件を再現して検証した。
