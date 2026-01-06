# Student Scheduler - 学生講座配置プログラム

学生の希望アンケートに基づいて、講座への配置を自動的に最適化するプログラムです。

## 機能

- Excel形式でのアンケート入力テンプレート生成
- 貪欲法 + 焼きなまし法による最適配置アルゴリズム
- 時限間の人数バランス調整
- 希望順位を考慮した配置最適化
- Excel形式での結果出力（サマリー、時限別、生徒別）

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
2. 生徒数と希望順位の数を入力
3. 自動生成されるExcelテンプレートにアンケートデータを入力
4. ファイルを保存して閉じる
5. 自動的に最適配置が計算され、結果がExcelファイルに出力されます

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
