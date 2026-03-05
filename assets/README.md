# Report Assets

`useLocalChartJs=true` を有効にした場合、以下の相対パスに Chart.js を配置してください。

- `assets/chart.umd.min.js`

`Run_PC_Optimizer.bat` / `PC_Optimizer.ps1` 実行時に、上記ファイルが存在すれば `reports/assets/` へコピーして
HTMLレポートからローカル参照します。

ファイルが無い場合は、レポート内で表形式のフォールバック表示に自動切替されます。

