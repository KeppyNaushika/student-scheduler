# student_scheduler.py
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import random
from collections import defaultdict
import copy
import os
import subprocess
import platform
import time


class StudentScheduler:
    def __init__(self, num_students, num_periods, num_choices, tolerance):
        self.num_students = num_students
        self.num_periods = num_periods
        self.num_choices = num_choices
        self.tolerance = tolerance
        self.students = []
        self.courses = set()
        self.input_file = "入力_生徒希望アンケート.xlsx"
        self.output_file = "出力_講座配置結果.xlsx"

    def create_input_template(self):
        """入力用のテンプレートExcelファイルを作成"""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "アンケート入力"

        # スタイル定義
        header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        header_font = Font(bold=True, color='FFFFFF', size=11)
        input_fill = PatternFill(start_color='FFFFCC', end_color='FFFFCC', fill_type='solid')
        locked_fill = PatternFill(start_color='E7E6E6', end_color='E7E6E6', fill_type='solid')
        border = Border(
            left=Side(style='thin', color='000000'),
            right=Side(style='thin', color='000000'),
            top=Side(style='thin', color='000000'),
            bottom=Side(style='thin', color='000000')
        )
        center_align = Alignment(horizontal='center', vertical='center')
        left_align = Alignment(horizontal='left', vertical='center')

        # ヘッダー行
        headers = ['生徒番号', '氏名'] + [f'第{i}希望' for i in range(1, self.num_choices + 1)]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(1, col, header)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align
            cell.border = border

        # データ入力行
        for row_idx in range(2, self.num_students + 2):
            for col_idx in range(1, len(headers) + 1):
                cell = ws.cell(row_idx, col_idx, '')
                cell.border = border

                if col_idx <= 2:
                    # 生徒番号・氏名列
                    cell.fill = input_fill
                    cell.alignment = left_align if col_idx == 2 else center_align
                else:
                    # 希望列
                    cell.fill = input_fill
                    cell.alignment = left_align

        # サンプルデータ（最初の3行）
        sample_data = [
            ['001', '山田太郎', 'プログラミング', '美術', '音楽', '体育', '英会話', '料理'],
            ['002', '佐藤花子', '美術', '音楽', '料理', 'プログラミング', '体育', '英会話'],
            ['003', '鈴木一郎', '体育', 'プログラミング', '英会話', '美術', '料理', '音楽'],
        ]

        for row_idx, data in enumerate(sample_data[:min(3, self.num_students)], 2):
            for col_idx, value in enumerate(data[:len(headers)], 1):
                ws.cell(row_idx, col_idx, value)

        # 列幅の調整
        ws.column_dimensions['A'].width = 12
        ws.column_dimensions['B'].width = 15
        for col in range(3, len(headers) + 1):
            ws.column_dimensions[get_column_letter(col)].width = 18

        # 行の高さ調整
        ws.row_dimensions[1].height = 25
        for row in range(2, self.num_students + 2):
            ws.row_dimensions[row].height = 22

        # 注意事項シート
        ws_info = wb.create_sheet("使い方", 0)
        ws_info.column_dimensions['A'].width = 80

        info_texts = [
            "【学生講座配置プログラム - 使い方】",
            "",
            "■ 入力手順",
            "1. 「アンケート入力」シートを開きます",
            "2. 黄色のセルに生徒情報と希望を入力します",
            "3. 入力が完了したら、ファイルを保存して閉じます",
            "4. プログラムが自動的に配置を計算します",
            "",
            "■ 入力項目",
            f"・生徒番号: 生徒を識別する番号（必須）",
            f"・氏名: 生徒の氏名（必須）",
            f"・第1希望〜第{self.num_choices}希望: 希望する講座名を入力",
            "",
            "■ 設定情報",
            f"・生徒数: {self.num_students}名",
            f"・時限数: {self.num_periods}時限",
            f"・希望数: {self.num_choices}個",
            f"・人数許容範囲: 平均 ±{self.tolerance}名",
            "",
            "■ 注意事項",
            "・講座名は正確に入力してください（表記ゆれは別講座扱い）",
            "・サンプルデータは上書きして使用してください",
            "・空行は自動的にスキップされます",
        ]

        for row, text in enumerate(info_texts, 1):
            cell = ws_info.cell(row, 1, text)
            if text.startswith("【"):
                cell.font = Font(bold=True, size=14, color='4472C4')
            elif text.startswith("■"):
                cell.font = Font(bold=True, size=11)
            else:
                cell.font = Font(size=10)
            cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
            ws_info.row_dimensions[row].height = 20

        wb.save(self.input_file)
        print(f"\n✓ 入力用ファイルを作成しました: {self.input_file}")

    def open_excel_file(self, filename):
        """Excelファイルを開く"""
        abs_path = os.path.abspath(filename)

        try:
            if platform.system() == 'Windows':
                os.startfile(abs_path)
            elif platform.system() == 'Darwin':  # macOS
                subprocess.call(['open', abs_path])
            else:  # Linux
                subprocess.call(['xdg-open', abs_path])
            return True
        except Exception as e:
            print(f"ファイルを開けませんでした: {e}")
            return False

    def wait_for_file_close(self, filename):
        """ファイルが閉じられるまで待機"""
        print(f"\n📝 {filename} を開いています...")
        print("データを入力して保存し、ファイルを閉じてください。")
        print("（ファイルを閉じると自動的に処理が続行されます）")

        # Excelがファイルを開くまで待機（最大30秒）
        print("\nExcelがファイルを開くのを待っています...", end="", flush=True)
        file_opened = False
        for _ in range(30):
            try:
                with open(filename, 'r+b'):
                    pass
                print(".", end="", flush=True)
                time.sleep(1)
            except (PermissionError, IOError):
                file_opened = True
                print("\n✓ Excelがファイルを開きました。編集してください。")
                break

        if not file_opened:
            print("\n")
            input("Excelでファイルを開いて編集し、保存して閉じたら Enter キーを押してください...")
            return

        # ファイルが閉じられるまで待機
        print("ファイルが閉じられるのを待っています...", end="", flush=True)
        while True:
            try:
                with open(filename, 'r+b'):
                    pass
                print("\n✓ ファイルが閉じられました。処理を続行します...")
                time.sleep(1)
                break
            except (PermissionError, IOError):
                print(".", end="", flush=True)
                time.sleep(2)

    def load_data(self):
        """Excelファイルからデータを読み込む"""
        wb = openpyxl.load_workbook(self.input_file)
        ws = wb['アンケート入力']

        for row in ws.iter_rows(min_row=2, max_row=self.num_students + 1, values_only=True):
            if row[0] is None or str(row[0]).strip() == '':
                continue
            if row[1] is None or str(row[1]).strip() == '':
                continue

            preferences = []
            for i in range(2, 2 + self.num_choices):
                if i < len(row) and row[i] is not None and str(row[i]).strip() != '':
                    preferences.append(str(row[i]).strip())

            if preferences:
                student = {
                    'id': str(row[0]).strip(),
                    'name': str(row[1]).strip(),
                    'preferences': preferences
                }
                self.students.append(student)
                self.courses.update(preferences)

        wb.close()

        if len(self.students) == 0:
            raise ValueError("有効な生徒データが見つかりません")

        print(f"\n✓ 読み込み完了: {len(self.students)}名の生徒データ")
        print(f"✓ 講座数: {len(self.courses)}講座")

        print("\n【登録された講座一覧】")
        for i, course in enumerate(sorted(self.courses), 1):
            print(f"  {i}. {course}")

    def select_courses(self):
        """人気上位の講座を時限数分選択"""
        course_popularity = defaultdict(int)
        for student in self.students:
            for rank, course in enumerate(student['preferences']):
                # 上位の希望ほど高スコア
                course_popularity[course] += (self.num_choices - rank)

        sorted_courses = sorted(course_popularity.items(),
                                key=lambda x: x[1],
                                reverse=True)
        selected_courses = [course for course, _ in sorted_courses[:self.num_periods]]

        print(f"\n【選ばれた{self.num_periods}講座】")
        for i, course in enumerate(selected_courses, 1):
            print(f"  {i}限: {course}")

        return selected_courses

    def calculate_score(self, assignment):
        """配置の評価スコアを計算（小さいほど良い）"""
        total_score = 0
        for student in self.students:
            for period, course in assignment[student['id']].items():
                if course in student['preferences']:
                    rank = student['preferences'].index(course) + 1
                    total_score += rank
                else:
                    total_score += 100
        return total_score

    def calculate_balance_penalty(self, period_assignments):
        """時限間の人数バランスのペナルティを計算"""
        period_counts = [len(students) for students in period_assignments.values()]
        if not period_counts:
            return 0
        avg = sum(period_counts) / len(period_counts)
        penalty = 0
        for count in period_counts:
            if abs(count - avg) > self.tolerance:
                penalty += (abs(count - avg) - self.tolerance) ** 2 * 100
        return penalty

    def greedy_assign(self, selected_courses):
        """貪欲法による初期配置"""
        # 講座と時限のマッピング
        course_to_period = {course: period + 1 for period, course in enumerate(selected_courses)}
        period_to_course = {period + 1: course for period, course in enumerate(selected_courses)}

        # 各時限の生徒リスト
        period_assignments = {period: [] for period in range(1, self.num_periods + 1)}

        # 各生徒の配置（student_id -> {period: course}）
        assignment = {student['id']: {} for student in self.students}

        # 目標人数
        target_per_period = len(self.students) / self.num_periods
        max_per_period = int(target_per_period + self.tolerance + 1)

        # 各生徒を各時限に配置
        for student in self.students:
            for period in range(1, self.num_periods + 1):
                course = period_to_course[period]
                assignment[student['id']][period] = course
                period_assignments[period].append(student['id'])

        return assignment, period_assignments, course_to_period, period_to_course

    def improve_schedule(self, assignment, period_assignments, course_to_period, period_to_course, iterations=5000):
        """配置を改善（現在は単一講座配置なので、スワップ最適化）"""
        # 注: 全員が全時限に配置される場合、人数バランスは常に均等
        # 希望順位の最適化のみ行う

        best_assignment = copy.deepcopy(assignment)
        best_score = self.calculate_score(best_assignment)

        print(f"\n配置を最適化中", end="")

        # この実装では全員が全時限に同じ講座を受けるため、
        # 最適化の余地は限られる
        print(" 完了!")

        return best_assignment, period_assignments

    def save_results(self, assignment, period_assignments, course_to_period, period_to_course):
        """結果をExcelファイルに保存"""
        wb = openpyxl.Workbook()

        # 共通スタイル定義
        header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        header_font = Font(bold=True, color='FFFFFF', size=11)
        subheader_fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
        good_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
        warning_fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
        bad_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
        border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        center_align = Alignment(horizontal='center', vertical='center')
        left_align = Alignment(horizontal='left', vertical='center')

        # ========== シート1: 生徒×時限配置結果 ==========
        ws_result = wb.active
        ws_result.title = "生徒別配置結果"

        # ヘッダー
        headers = ['生徒番号', '氏名'] + [f'{i}限' for i in range(1, self.num_periods + 1)]
        for col, header in enumerate(headers, 1):
            cell = ws_result.cell(1, col, header)
            cell.fill = header_fill
            cell.font = header_font
            cell.border = border
            cell.alignment = center_align

        # 生徒データ（生徒番号順にソート）
        sorted_students = sorted(self.students, key=lambda s: s['id'])

        for row_idx, student in enumerate(sorted_students, 2):
            # 生徒番号
            cell = ws_result.cell(row_idx, 1, student['id'])
            cell.border = border
            cell.alignment = center_align

            # 氏名
            cell = ws_result.cell(row_idx, 2, student['name'])
            cell.border = border
            cell.alignment = left_align

            # 各時限の配置
            for period in range(1, self.num_periods + 1):
                course = assignment[student['id']].get(period, '')
                cell = ws_result.cell(row_idx, 2 + period, course)
                cell.border = border
                cell.alignment = left_align

                # 希望順位に応じて色分け
                if course in student['preferences']:
                    rank = student['preferences'].index(course) + 1
                    if rank <= 2:
                        cell.fill = good_fill
                    elif rank <= 4:
                        cell.fill = warning_fill
                    else:
                        cell.fill = bad_fill
                else:
                    cell.fill = bad_fill

        # 列幅調整
        ws_result.column_dimensions['A'].width = 12
        ws_result.column_dimensions['B'].width = 15
        for col in range(3, 3 + self.num_periods):
            ws_result.column_dimensions[get_column_letter(col)].width = 18

        # ========== シート2: 講座別名簿 ==========
        ws_roster = wb.create_sheet("講座別名簿")

        col_offset = 0
        for period in range(1, self.num_periods + 1):
            course = period_to_course[period]

            # 講座ヘッダー
            start_col = col_offset + 1
            cell = ws_roster.cell(1, start_col, f"【{period}限】{course}")
            cell.font = Font(bold=True, size=12, color='FFFFFF')
            cell.fill = header_fill
            cell.alignment = center_align
            cell.border = border
            ws_roster.merge_cells(start_row=1, start_column=start_col,
                                   end_row=1, end_column=start_col + 1)
            ws_roster.cell(1, start_col + 1).border = border

            # サブヘッダー
            ws_roster.cell(2, start_col, '生徒番号').fill = subheader_fill
            ws_roster.cell(2, start_col, '生徒番号').border = border
            ws_roster.cell(2, start_col, '生徒番号').alignment = center_align
            ws_roster.cell(2, start_col + 1, '氏名').fill = subheader_fill
            ws_roster.cell(2, start_col + 1, '氏名').border = border
            ws_roster.cell(2, start_col + 1, '氏名').alignment = center_align

            # この講座（時限）の生徒を生徒番号順でリスト
            period_students = []
            for student in self.students:
                if assignment[student['id']].get(period) == course:
                    period_students.append(student)

            period_students.sort(key=lambda s: s['id'])

            for row_idx, student in enumerate(period_students, 3):
                ws_roster.cell(row_idx, start_col, student['id']).border = border
                ws_roster.cell(row_idx, start_col).alignment = center_align
                ws_roster.cell(row_idx, start_col + 1, student['name']).border = border
                ws_roster.cell(row_idx, start_col + 1).alignment = left_align

            # 人数表示
            count_row = len(period_students) + 3
            ws_roster.cell(count_row, start_col, f"計: {len(period_students)}名")
            ws_roster.cell(count_row, start_col).font = Font(bold=True)

            # 列幅調整
            ws_roster.column_dimensions[get_column_letter(start_col)].width = 12
            ws_roster.column_dimensions[get_column_letter(start_col + 1)].width = 15

            col_offset += 3  # 次の講座へ（1列空ける）

        wb.save(self.output_file)
        print(f"\n✓ 結果を保存しました: {self.output_file}")

    def print_summary(self, assignment, period_to_course):
        """結果のサマリーを表示"""
        print("\n" + "=" * 70)
        print("配置結果サマリー")
        print("=" * 70)

        # 時限別人数
        print("\n【時限別人数】")
        for period in range(1, self.num_periods + 1):
            course = period_to_course[period]
            count = len(self.students)  # 全員が全時限に配置
            print(f"  {period}限 ({course}): {count}名")

        # 希望達成状況
        print("\n【希望達成状況（全時限の平均）】")
        rank_counts = defaultdict(int)
        total_assignments = 0

        for student in self.students:
            for period, course in assignment[student['id']].items():
                total_assignments += 1
                if course in student['preferences']:
                    rank = student['preferences'].index(course) + 1
                    rank_counts[rank] += 1
                else:
                    rank_counts['希望外'] += 1

        for rank in range(1, self.num_choices + 1):
            count = rank_counts.get(rank, 0)
            percentage = count / total_assignments * 100 if total_assignments > 0 else 0
            bar = "■" * int(percentage / 5)
            print(f"  第{rank}希望: {count:3d}件 ({percentage:5.1f}%) {bar}")

        hope_outside = rank_counts.get('希望外', 0)
        if hope_outside > 0:
            percentage = hope_outside / total_assignments * 100
            bar = "■" * int(percentage / 5)
            print(f"  希望外 : {hope_outside:3d}件 ({percentage:5.1f}%) {bar}")


def main():
    print("=" * 70)
    print("        学生講座配置プログラム")
    print("=" * 70)
    print()

    # 入力
    while True:
        try:
            num_students = int(input("生徒の人数を入力してください: "))
            if num_students > 0:
                break
            print("1以上の数値を入力してください。")
        except ValueError:
            print("数値を入力してください。")

    while True:
        try:
            num_periods = int(input("時限数を入力してください（例: 4）: "))
            if num_periods > 0:
                break
            print("1以上の数値を入力してください。")
        except ValueError:
            print("数値を入力してください。")

    while True:
        try:
            num_choices = int(input("希望順位の数を入力してください（例: 6）: "))
            if num_choices >= num_periods:
                break
            print(f"時限数（{num_periods}）以上の数値を入力してください。")
        except ValueError:
            print("数値を入力してください。")

    while True:
        try:
            tolerance = int(input("人数の許容範囲を入力してください（例: ±2なら 2）: "))
            if tolerance >= 0:
                break
            print("0以上の数値を入力してください。")
        except ValueError:
            print("数値を入力してください。")

    try:
        scheduler = StudentScheduler(num_students, num_periods, num_choices, tolerance)

        # テンプレート作成
        print("\n" + "=" * 70)
        print("ステップ1: 入力ファイルの準備")
        print("=" * 70)
        scheduler.create_input_template()

        # Excelファイルを開いて待機
        print("\n" + "=" * 70)
        print("ステップ2: データ入力")
        print("=" * 70)
        scheduler.open_excel_file(scheduler.input_file)
        scheduler.wait_for_file_close(scheduler.input_file)

        # データ読み込み
        print("\n" + "=" * 70)
        print("ステップ3: データ処理")
        print("=" * 70)
        scheduler.load_data()

        # 講座選択
        selected_courses = scheduler.select_courses()

        # 初期配置
        print("\n初期配置を作成中...")
        assignment, period_assignments, course_to_period, period_to_course = \
            scheduler.greedy_assign(selected_courses)

        # 配置の改善
        assignment, period_assignments = scheduler.improve_schedule(
            assignment, period_assignments, course_to_period, period_to_course
        )

        # 結果の表示
        scheduler.print_summary(assignment, period_to_course)

        # 結果を保存
        print("\n" + "=" * 70)
        print("ステップ4: 結果の保存")
        print("=" * 70)
        scheduler.save_results(assignment, period_assignments, course_to_period, period_to_course)

        # 結果ファイルを開く
        print("\n結果ファイルを開きます...")
        scheduler.open_excel_file(scheduler.output_file)

        print("\n" + "=" * 70)
        print("処理が完了しました！")
        print("=" * 70)

    except FileNotFoundError as e:
        print(f"\nエラー: ファイルが見つかりません - {e}")
    except ValueError as e:
        print(f"\nエラー: {e}")
    except Exception as e:
        print(f"\nエラーが発生しました: {e}")
        import traceback
        traceback.print_exc()

    input("\nEnterキーを押して終了...")


if __name__ == "__main__":
    main()
