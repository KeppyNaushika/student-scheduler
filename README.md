# Student Scheduler - 学生講座配置プログラム

学生の希望アンケートに基づいて、講座への配置を自動的に最適化するプログラムです。

## 機能

- Excel形式でのアンケート入力（生徒番号、氏名、希望講座）
- 人気上位の講座を自動選択
- 全時限への配置と結果出力
- 人数バランスの許容範囲設定
- 2種類の出力シート
  - **生徒別配置結果**: 生徒×時限の配置表
  - **講座別名簿**: 各講座の生徒番号順名簿

## インストール

### 方法1: Windows実行ファイル（推奨）

[Releases](../../releases) ページから `student-scheduler.exe` をダウンロードして実行してください。

### 方法2: Python環境で実行

```bash
# uv を使用
uv sync
uv run python main.py

# または pip を使用
pip install openpyxl
python main.py
```

## 使い方

1. プログラムを起動
2. 以下の情報を入力:
   - 生徒数
   - 時限数
   - 希望順位の数
   - 人数の許容範囲（±N人）
3. 自動生成されるExcelテンプレートにアンケートデータを入力
   - 黄色のセルに入力
   - 生徒番号、氏名、第1希望〜第N希望
4. ファイルを保存して閉じる
5. 自動的に配置が計算され、結果がExcelファイルに出力されます

## 入出力ファイル

| ファイル | 説明 |
|----------|------|
| `入力_生徒希望アンケート.xlsx` | 生徒情報と希望講座の入力用 |
| `出力_講座配置結果.xlsx` | 配置結果（2シート） |

## 開発

```bash
# 依存関係のインストール
uv sync --group dev

# 実行
uv run python main.py

# EXE作成（Windows環境）
uv run pyinstaller --onefile --name student-scheduler main.py
```

## ライセンス

MIT License
