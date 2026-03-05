
高性能PC最適化ツール詳細要件定義書（完全版）

Version: 4.0
Project: PC_Optimizer
対象環境: 社内IT部門専用
対象機器: 社内Windows PC

---

# 1. システム概要

本システムは社内IT部門が管理するWindows PCを対象に、

* PC状態診断
* システム最適化
* セキュリティ確認
* Windows更新管理
* パフォーマンス分析
* HTMLレポート生成

を自動実行する **PowerShellベースのPC管理ツール**である。

---

# 2. システム目的

本ツールの目的

### PC運用効率向上

* PC状態の可視化
* 障害予防
* メンテナンス自動化

### IT部門運用支援

* PC診断レポート
* PC健康スコア
* 改善提案提示

---

# 3. 利用対象

利用者

```
社内IT部門スタッフ
```

対象機器

```
社内Windows PC
```

外部公開

```
なし（社内専用）
```

---

# 4. 対応環境

| 項目         | 内容                      |
| ---------- | ----------------------- |
| OS         | Windows 10 / Windows 11 |
| PowerShell | 5.1 / 7.x               |
| 権限         | 管理者権限                   |
| ネットワーク     | 社内LAN                   |

---

# 5. システム構成

システム構造

```
Run_PC_Optimizer.bat
        ↓
PC_Optimizer.ps1
        ↓
Modules
        ↓
処理実行
        ↓
ログ + HTMLレポート生成
```

---

# 6. GitHubプロジェクト構成（ベスト構成）

推奨ディレクトリ構造

```
PC_Optimizer
│
├─ PC_Optimizer.ps1
├─ Run_PC_Optimizer.bat
├─ CLAUDE.md
│
├─ config
│   └─ config.json
│
├─ modules
│   ├─ Cleanup.psm1
│   ├─ Diagnostics.psm1
│   ├─ Performance.psm1
│   ├─ Network.psm1
│   ├─ Security.psm1
│   ├─ Update.psm1
│   └─ Report.psm1
│
├─ reports
│
├─ logs
│
├─ docs
│
├─ tests
│
└─ .github
    └─ workflows
```

---

# 7. Git管理方針

`.gitignore`

```
logs/*
reports/*
.env
*.log
*.tmp
```

---

# 8. config仕様

設定ファイル

```
config/config.json
```

例

```json
{
 "logLevel": "INFO",
 "reportFormat": "HTML",
 "cleanupBrowserCache": true,
 "diskOptimize": true,
 "enableSecurityCheck": true
}
```

---

# 9. 機能一覧

本ツールは以下の機能を提供する。

---

# 10. 機能① システム診断

取得情報

* CPU
* メモリ
* ディスク
* OS
* GPU

目的

```
PC状態診断
```

---

# 11. 機能② パフォーマンス分析

取得情報

* CPU負荷ランキング
* メモリ使用プロセス
* ディスクI/O

---

# 12. 機能③ スタートアップ分析

取得

* 自動起動アプリ
* スタートアップ数

評価

| 数    | 評価 |
| ---- | -- |
| 5以下  | 良  |
| 6〜10 | 普通 |
| 11以上 | 多  |

---

# 13. 機能④ セキュリティ診断

取得

* Windows Defender
* Firewall
* BitLocker
* UAC

---

# 14. 機能⑤ ネットワーク診断

取得

* IPアドレス
* DNS
* Gateway
* NIC速度

---

# 15. 機能⑥ Microsoft365診断

取得

* OneDrive同期状態
* Teamsキャッシュ
* Outlook状態

---

# 16. 機能⑦ Windows Update診断

取得

* 更新状態
* 更新履歴
* 更新エラー

---

# 17. 機能⑧ イベントログ分析

取得

* Application Error
* System Error
* BSOD履歴

---

# 18. 機能⑨ IT資産情報収集

取得

```
PC名
ユーザー
OS
IP
MAC
CPU
RAM
Disk
```

---

# 19. 機能⑩ レポート生成

出力形式

```
HTML
CSV
JSON
```

---

# 20. 機能⑪ GUIモード（将来）

PowerShell GUI

```
Windows Forms
WPF
```

---

# 21. 機能⑫ 定期メンテナンス

Task Scheduler

例

```
月1回
```

---

# 22. 機能⑬ 複数PC診断（将来）

技術

```
WinRM
PowerShell Remoting
```

---

# 23. 機能⑭ 自動修復

修復機能

```
DNS修復
Update修復
Temp削除
```

---

# 24. 機能⑮ AI診断（簡易版）

AI未使用

```
ルールベース診断
```

---

# 25. HTMLレポート仕様

レポート構造

```
PC Health Report
│
├─ 基本情報
├─ 総合評価
├─ CPU評価
├─ メモリ評価
├─ ディスク評価
├─ セキュリティ評価
├─ ネットワーク評価
├─ スタートアップ評価
├─ ログ分析
├─ 改善提案
└─ 実行ログ
```

---

# 26. 総合評価スコア

PC健康スコア

```
0〜100
```

評価

| スコア    | 評価        |
| ------ | --------- |
| 90〜100 | Excellent |
| 75〜89  | Good      |
| 60〜74  | Warning   |
| 59以下   | Critical  |

---

# 27. スコア計算

| 項目             | 配点 |
| -------------- | -- |
| CPU            | 15 |
| Memory         | 15 |
| Disk           | 20 |
| Startup        | 10 |
| Security       | 15 |
| Network        | 10 |
| Windows Update | 10 |
| System Health  | 5  |

---

# 28. 改善提案

例

```
スタートアップが多すぎます
ディスク空き容量不足
Windows Update未適用
```

---

# 29. HTML UI

レポート例

```
PC Health Report
----------------------------

PC Name: PC-001
User: user1

Score: 87 / 100

CPU       OK
Memory    OK
Disk      Warning
Startup   Warning
Security  OK
Network   OK
```

---

# 30. グラフ表示

表示

```
CPU使用率
メモリ使用率
ディスク使用率
```

JS

```
Chart.js
```

---

# 31. ログ仕様

ログ例

```
[2026-03-04 10:25:01] [INFO] Temp files deleted
```

---

# 32. セキュリティ

対策

* 管理者権限確認
* 削除パス制御
* 誤削除防止
* APIキー管理(.env)

---

# 33. PowerShell規約

命名規則

```
Verb-Noun
```

例

```
Get-SystemInfo
Clear-TempFiles
Test-NetworkConnection
```

---

# 34. エラーハンドリング

```
try
catch
finally
```

---

# 35. テスト

フレームワーク

```
Pester
```

---

# 36. CI

GitHub Actions

```
.github/workflows/ci.yml
```

機能

* PowerShellテスト
* PR検証

---

# 37. バージョン管理

方式

```
Semantic Versioning
```

例

```
v4.0.0
```

---

# 38. 将来拡張

拡張予定

```
GUI
WebUI
AI診断
IT資産管理
ダッシュボード
```

---

# まとめ

本ツールは

```
PCクリーナー
↓
PC最適化ツール
↓
PC診断ツール
↓
IT運用ツール
```

へ発展可能な設計とする。

---

