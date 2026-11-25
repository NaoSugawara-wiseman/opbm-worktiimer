# OBPM 工数入力の自動化ツール（作業タイマー＋自動入力）

ノート PC の作業タイマーで作業時間を記録し、「本日終了」を押すと、デスクトップ PC が SI Object Browser PM（OBPM）の工数入力画面に自動で入力・更新するツールです。Windows 標準の PowerShell 5.1 だけで動きます。

- 使い方・設定方法（利用者向け）: [docs/README_利用者向け.md](docs/README_利用者向け.md)
- 配布用の zip: [Releases](../../releases) からダウンロード

## フォルダ構成

| フォルダ | 中身 |
| --- | --- |
| `laptop/` | ノート PC 用の作業タイマー（`worktimer.bat` / `worktimer.ps1`） |
| `desktop/` | デスクトップ PC 用の自動入力（`obpm_input.*`）と画面調査ツール（`obpm_inspect.*`） |
| `common/` | 両方の PC で使う共通部品（`common.ps1`）。配布時に両方のフォルダへコピーする |
| `config/` | 設定ファイルの見本（`*.sample`）。配布時に `.sample` を外した名前で入る |
| `docs/` | 利用者向けの説明書 |
| `tools/` | 配布用 zip を作るスクリプト（`build.ps1`） |
| `VERSION` | 版番号 |

## 配布用 zip の作り方

リポジトリのフォルダで次を実行すると、`dist\worktimer_powershell_v<VERSION>.zip` ができます（「ノートPC用」「デスクトップPC用」フォルダと README の形）。

```powershell
powershell -ExecutionPolicy Bypass -File tools\build.ps1
```

## リリースの手順

1. 修正はブランチ（`fix/〇〇`・`feature/〇〇`）で行い、確認できたら `main` に入れる
2. `VERSION` を上げる（不具合修正 = 3 桁目、機能追加 = 2 桁目）
3. `tools\build.ps1` で zip を作る
4. GitHub でタグ `v<VERSION>` の Release を作り、zip を添付する。直した内容を数行書く

## 編集するときの注意（動かなくなるのを防ぐ）

| ファイル | 文字コード | 改行 |
| --- | --- | --- |
| `*.bat` | **Shift_JIS（ANSI）** | CRLF |
| `*.ps1` | **UTF-8（BOM 付き）** | CRLF |
| `*.ini.sample` / `*.csv.sample` / `*.md` | UTF-8 | CRLF |

- 文字コードと改行は `.gitattributes` で固定しています。Windows の Git では `core.autocrlf` を `false` にしておくと確実です（`git config core.autocrlf false`）
- PowerShell 5.1 で動くこと（`Set-StrictMode -Version 2.0`）を前提にしています。PowerShell 7 だけの書き方は使わない

## リポジトリに入れないもの

`.gitignore` で除外しています。誤って登録しないよう注意してください。

- 自動ログインの情報（`obpm_login.dat`：暗号化したパスワード）とその関連ファイル
- 実際に使っている設定（`*.ini`・`tasks.csv`）
- 共有フォルダの中身（作業記録・ログ・状態のファイル）
