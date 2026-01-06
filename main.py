# student_scheduler.py
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from collections import defaultdict
import os
import sys
import subprocess
import platform
import time

# PuLP for Integer Linear Programming
from pulp import (
    LpProblem, LpMinimize, LpVariable, LpBinary, lpSum, LpStatus, value,
    COIN_CMD
)

def get_solver():
    """ソルバーを取得（PyInstallerバンドル時はパスを指定）"""
    if getattr(sys, 'frozen', False):
        # PyInstallerでバンドルされた場合
        base_path = sys._MEIPASS  # type: ignore[attr-defined]
        if platform.system() == 'Windows':
            cbc_path = os.path.join(base_path, 'pulp', 'solverdir', 'cbc', 'win', '64', 'cbc.exe')
        elif platform.system() == 'Darwin':
            cbc_path = os.path.join(base_path, 'pulp', 'solverdir', 'cbc', 'osx', '64', 'cbc')
        else:
            cbc_path = os.path.join(base_path, 'pulp', 'solverdir', 'cbc', 'linux', 'i64', 'cbc')
        return COIN_CMD(path=cbc_path, msg=0)
    return COIN_CMD(msg=0)


class StudentScheduler:
    def __init__(self, num_students, num_periods, num_choices, min_per_course, max_per_course):
        self.num_students = num_students
        self.num_periods = num_periods  # 受講する講座数（例: 4）
        self.num_choices = num_choices  # 希望順位の数（例: 6、これが講座数）
        self.min_per_course = min_per_course
        self.max_per_course = max_per_course
        self.students = []
        self.courses = []  # 全講座リスト
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
                cell.fill = input_fill
                cell.alignment = left_align if col_idx == 2 else center_align

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
            f"・講座数: {self.num_choices}講座（全講座が開講されます）",
            f"・受講数: 各生徒は{self.num_periods}講座を受講",
            f"・時限数: {self.num_periods}時限",
            f"・人数範囲: {self.min_per_course}〜{self.max_per_course}名/講座/時限",
            "",
            "■ 配置ルール",
            f"・各時限で全{self.num_choices}講座が開講されます",
            f"・各生徒は{self.num_choices}講座のうち{self.num_periods}講座を受講します",
            "・生徒によって受講する講座の組み合わせは異なります",
            "・整数線形計画法(ILP)により最適解を計算します",
            "・各講座の人数ができるだけ均等になるよう調整されます",
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

        all_courses = set()

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
                all_courses.update(preferences)

        wb.close()

        if len(self.students) == 0:
            raise ValueError("有効な生徒データが見つかりません")

        # 全講座をリスト化（希望順位のスコアでソート）
        course_scores = defaultdict(int)
        for student in self.students:
            for rank, course in enumerate(student['preferences']):
                course_scores[course] += (self.num_choices - rank)

        self.courses = sorted(all_courses, key=lambda c: course_scores[c], reverse=True)

        print(f"\n✓ 読み込み完了: {len(self.students)}名の生徒データ")
        print(f"✓ 講座数: {len(self.courses)}講座")

        print("\n【講座一覧】（人気順）")
        for i, course in enumerate(self.courses, 1):
            print(f"  {i}. {course} (スコア: {course_scores[course]})")

    def get_preference_rank(self, student, course):
        """生徒の希望順位を取得（1始まり、希望外は大きな値）"""
        if course in student['preferences']:
            return student['preferences'].index(course) + 1
        return self.num_choices + 1  # 希望外

    def solve_with_ilp(self):
        """
        整数線形計画法(ILP)で最適配置を求める

        決定変数:
            x[s,c,p] = 1 if student s takes course c in period p

        目的関数:
            minimize Σ (preference_rank[s,c] * x[s,c,p]) + fairness_penalty

        制約:
            1. 各生徒は各時限で1つの講座を受講
            2. 各生徒は各講座を最大1回受講
            3. 各生徒はnum_periods個の講座を受講
            4. 各時限の各講座の人数は目標±許容範囲
        """
        print("\n【整数線形計画法(ILP)で最適化】")
        print("問題を定式化中...")

        # 問題の作成
        prob = LpProblem("StudentScheduler", LpMinimize)

        # インデックス
        students_idx = range(len(self.students))
        courses_idx = range(len(self.courses))
        periods_idx = range(1, self.num_periods + 1)

        # 決定変数: x[s][c][p] = 1 if student s takes course c in period p
        x = {}
        for s in students_idx:
            for c in courses_idx:
                for p in periods_idx:
                    x[s, c, p] = LpVariable(f"x_{s}_{c}_{p}", cat=LpBinary)

        # 補助変数: y[s][c] = 1 if student s takes course c (any period)
        y = {}
        for s in students_idx:
            for c in courses_idx:
                y[s, c] = LpVariable(f"y_{s}_{c}", cat=LpBinary)

        # 公平性のための補助変数
        max_score = LpVariable("max_score", lowBound=0)
        min_score = LpVariable("min_score", lowBound=0)

        # 各生徒のスコア（希望順位の合計）
        student_scores = {}
        for s in students_idx:
            student = self.students[s]
            student_scores[s] = lpSum(
                self.get_preference_rank(student, self.courses[c]) * y[s, c]
                for c in courses_idx
            )

        print("目的関数を設定中...")

        # 目的関数: 希望順位の合計 + 公平性ペナルティ
        total_preference_score = lpSum(student_scores[s] for s in students_idx)
        fairness_penalty = (max_score - min_score) * 10

        prob += total_preference_score + fairness_penalty, "Total_Cost"

        print("制約条件を追加中...")

        # 制約1: 各生徒は各時限で1つの講座を受講
        for s in students_idx:
            for p in periods_idx:
                prob += lpSum(x[s, c, p] for c in courses_idx) == 1, f"OnePerPeriod_s{s}_p{p}"

        # 制約2: 各生徒は各講座を最大1回受講
        for s in students_idx:
            for c in courses_idx:
                prob += lpSum(x[s, c, p] for p in periods_idx) <= 1, f"MaxOnce_s{s}_c{c}"

        # 制約3: y[s,c]とx[s,c,p]の関係
        for s in students_idx:
            for c in courses_idx:
                prob += y[s, c] == lpSum(x[s, c, p] for p in periods_idx), f"Link_y_x_s{s}_c{c}"

        # 制約4: 各時限の各講座の人数バランス
        for p in periods_idx:
            for c in courses_idx:
                count = lpSum(x[s, c, p] for s in students_idx)
                prob += count >= self.min_per_course, f"MinBalance_p{p}_c{c}"
                prob += count <= self.max_per_course, f"MaxBalance_p{p}_c{c}"

        # 制約5: 公平性（max_score, min_score）
        for s in students_idx:
            prob += student_scores[s] <= max_score, f"MaxScore_s{s}"
            prob += student_scores[s] >= min_score, f"MinScore_s{s}"

        print(f"変数数: {len(prob.variables())}")
        print(f"制約数: {len(prob.constraints)}")
        print("\n最適化を実行中（しばらくお待ちください）...")

        # 求解
        start_time = time.time()
        prob.solve(get_solver())
        solve_time = time.time() - start_time

        print(f"\n✓ 求解完了（{solve_time:.1f}秒）")
        print(f"ステータス: {LpStatus[prob.status]}")

        if prob.status != 1:  # 1 = Optimal
            print("警告: 最適解が見つかりませんでした。制約を緩和して再試行します...")
            return self.solve_with_relaxed_constraints()

        # 結果の抽出
        course_selection = {student['id']: set() for student in self.students}
        schedule = {student['id']: {} for student in self.students}

        for s in students_idx:
            student_id = self.students[s]['id']
            for c in courses_idx:
                for p in periods_idx:
                    if value(x[s, c, p]) and value(x[s, c, p]) > 0.5:
                        course_name = self.courses[c]
                        course_selection[student_id].add(course_name)
                        schedule[student_id][p] = course_name

        # 目的関数の値
        print(f"目的関数値: {value(prob.objective):.2f}")

        return course_selection, schedule

    def solve_with_relaxed_constraints(self):
        """制約を緩和して解を求める（フォールバック）"""
        print("\n制約を緩和して再試行...")

        prob = LpProblem("StudentScheduler_Relaxed", LpMinimize)

        students_idx = range(len(self.students))
        courses_idx = range(len(self.courses))
        periods_idx = range(1, self.num_periods + 1)

        x = {}
        for s in students_idx:
            for c in courses_idx:
                for p in periods_idx:
                    x[s, c, p] = LpVariable(f"x_{s}_{c}_{p}", cat=LpBinary)

        # 目的関数（公平性ペナルティなし）
        prob += lpSum(
            self.get_preference_rank(self.students[s], self.courses[c]) * x[s, c, p]
            for s in students_idx
            for c in courses_idx
            for p in periods_idx
        ), "Total_Preference"

        # 制約1: 各生徒は各時限で1つの講座
        for s in students_idx:
            for p in periods_idx:
                prob += lpSum(x[s, c, p] for c in courses_idx) == 1

        # 制約2: 各生徒は各講座を最大1回
        for s in students_idx:
            for c in courses_idx:
                prob += lpSum(x[s, c, p] for p in periods_idx) <= 1

        # 制約3: 人数バランス（緩和）
        relaxed_min = max(0, self.min_per_course - 5)
        relaxed_max = self.max_per_course + 5
        for p in periods_idx:
            for c in courses_idx:
                count = lpSum(x[s, c, p] for s in students_idx)
                prob += count >= relaxed_min
                prob += count <= relaxed_max

        prob.solve(get_solver())

        if prob.status != 1:
            raise ValueError("最適化に失敗しました。入力データを確認してください。")

        course_selection = {student['id']: set() for student in self.students}
        schedule = {student['id']: {} for student in self.students}

        for s in students_idx:
            student_id = self.students[s]['id']
            for c in courses_idx:
                for p in periods_idx:
                    if value(x[s, c, p]) and value(x[s, c, p]) > 0.5:
                        course_name = self.courses[c]
                        course_selection[student_id].add(course_name)
                        schedule[student_id][p] = course_name

        return course_selection, schedule

    def save_results(self, course_selection, schedule):
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

        headers = ['生徒番号', '氏名'] + [f'{i}限' for i in range(1, self.num_periods + 1)]
        for col, header in enumerate(headers, 1):
            cell = ws_result.cell(1, col, header)
            cell.fill = header_fill
            cell.font = header_font
            cell.border = border
            cell.alignment = center_align

        sorted_students = sorted(self.students, key=lambda s: s['id'])

        for row_idx, student in enumerate(sorted_students, 2):
            cell = ws_result.cell(row_idx, 1, student['id'])
            cell.border = border
            cell.alignment = center_align

            cell = ws_result.cell(row_idx, 2, student['name'])
            cell.border = border
            cell.alignment = left_align

            for period in range(1, self.num_periods + 1):
                course = schedule[student['id']].get(period, '')
                cell = ws_result.cell(row_idx, 2 + period, course)
                cell.border = border
                cell.alignment = left_align

                rank = self.get_preference_rank(student, course)
                if rank <= 2:
                    cell.fill = good_fill
                elif rank <= 4:
                    cell.fill = warning_fill
                elif rank <= self.num_choices:
                    cell.fill = bad_fill

        ws_result.column_dimensions['A'].width = 12
        ws_result.column_dimensions['B'].width = 15
        for col in range(3, 3 + self.num_periods):
            ws_result.column_dimensions[get_column_letter(col)].width = 18

        # ========== シート2: 講座別名簿 ==========
        ws_roster = wb.create_sheet("講座別名簿")

        col_offset = 0
        for period in range(1, self.num_periods + 1):
            for course in self.courses:
                start_col = col_offset + 1
                cell = ws_roster.cell(1, start_col, f"【{period}限】{course}")
                cell.font = Font(bold=True, size=11, color='FFFFFF')
                cell.fill = header_fill
                cell.alignment = center_align
                cell.border = border
                ws_roster.merge_cells(start_row=1, start_column=start_col,
                                       end_row=1, end_column=start_col + 1)
                ws_roster.cell(1, start_col + 1).border = border

                ws_roster.cell(2, start_col, '生徒番号').fill = subheader_fill
                ws_roster.cell(2, start_col, '生徒番号').border = border
                ws_roster.cell(2, start_col, '生徒番号').alignment = center_align
                ws_roster.cell(2, start_col + 1, '氏名').fill = subheader_fill
                ws_roster.cell(2, start_col + 1, '氏名').border = border
                ws_roster.cell(2, start_col + 1, '氏名').alignment = center_align

                course_students = [s for s in self.students
                                   if schedule[s['id']].get(period) == course]
                course_students.sort(key=lambda s: s['id'])

                for row_idx, student in enumerate(course_students, 3):
                    ws_roster.cell(row_idx, start_col, student['id']).border = border
                    ws_roster.cell(row_idx, start_col).alignment = center_align
                    ws_roster.cell(row_idx, start_col + 1, student['name']).border = border
                    ws_roster.cell(row_idx, start_col + 1).alignment = left_align

                count_row = max(len(course_students) + 3, 4)
                ws_roster.cell(count_row, start_col, f"計: {len(course_students)}名")
                ws_roster.cell(count_row, start_col).font = Font(bold=True)

                ws_roster.column_dimensions[get_column_letter(start_col)].width = 10
                ws_roster.column_dimensions[get_column_letter(start_col + 1)].width = 12

                col_offset += 3

            col_offset += 1

        # ========== シート3: 希望達成度 ==========
        ws_stats = wb.create_sheet("希望達成度")

        stat_headers = ['生徒番号', '氏名', '満足度', '平均順位'] + \
                       [f'第{i}希望' for i in range(1, self.num_choices + 1)] + ['希望外']
        for col, header in enumerate(stat_headers, 1):
            cell = ws_stats.cell(1, col, header)
            cell.fill = header_fill
            cell.font = header_font
            cell.border = border
            cell.alignment = center_align

        student_stats = []
        for student in sorted_students:
            rank_counts = defaultdict(int)
            total_rank = 0
            count = 0

            selected = course_selection.get(student['id'], set())
            for course in selected:
                rank = self.get_preference_rank(student, course)
                if rank <= self.num_choices:
                    rank_counts[rank] += 1
                    total_rank += rank
                else:
                    rank_counts['希望外'] += 1
                    total_rank += self.num_choices + 1
                count += 1

            avg_rank = total_rank / count if count > 0 else 0
            max_possible = self.num_periods
            min_possible = self.num_periods * (self.num_choices + 1)
            satisfaction = 100 * (min_possible - total_rank) / (min_possible - max_possible) if min_possible > max_possible else 100

            student_stats.append({
                'student': student,
                'satisfaction': satisfaction,
                'avg_rank': avg_rank,
                'rank_counts': rank_counts
            })

        student_stats.sort(key=lambda x: x['satisfaction'])

        for row_idx, stat in enumerate(student_stats, 2):
            student = stat['student']

            cell = ws_stats.cell(row_idx, 1, student['id'])
            cell.border = border
            cell.alignment = center_align

            cell = ws_stats.cell(row_idx, 2, student['name'])
            cell.border = border
            cell.alignment = left_align

            cell = ws_stats.cell(row_idx, 3, round(stat['satisfaction'], 1))
            cell.border = border
            cell.alignment = center_align
            if stat['satisfaction'] >= 80:
                cell.fill = good_fill
            elif stat['satisfaction'] >= 60:
                cell.fill = warning_fill
            else:
                cell.fill = bad_fill

            cell = ws_stats.cell(row_idx, 4, round(stat['avg_rank'], 2))
            cell.border = border
            cell.alignment = center_align

            for rank in range(1, self.num_choices + 1):
                cell = ws_stats.cell(row_idx, 4 + rank, stat['rank_counts'].get(rank, 0))
                cell.border = border
                cell.alignment = center_align

            cell = ws_stats.cell(row_idx, 5 + self.num_choices, stat['rank_counts'].get('希望外', 0))
            cell.border = border
            cell.alignment = center_align

        summary_row = len(student_stats) + 3
        ws_stats.cell(summary_row, 1, '【統計】').font = Font(bold=True)

        satisfactions = [s['satisfaction'] for s in student_stats]
        avg_ranks = [s['avg_rank'] for s in student_stats]

        stats_info = [
            (summary_row + 1, '平均満足度', f"{sum(satisfactions)/len(satisfactions):.1f}点"),
            (summary_row + 2, '最低満足度', f"{min(satisfactions):.1f}点"),
            (summary_row + 3, '最高満足度', f"{max(satisfactions):.1f}点"),
            (summary_row + 4, '標準偏差', f"{(sum((s-sum(satisfactions)/len(satisfactions))**2 for s in satisfactions)/len(satisfactions))**0.5:.2f}"),
            (summary_row + 5, '平均希望順位', f"{sum(avg_ranks)/len(avg_ranks):.2f}"),
        ]

        for row, label, val in stats_info:
            ws_stats.cell(row, 1, label).font = Font(bold=True)
            ws_stats.cell(row, 2, val)

        ws_stats.column_dimensions['A'].width = 12
        ws_stats.column_dimensions['B'].width = 15
        ws_stats.column_dimensions['C'].width = 12
        ws_stats.column_dimensions['D'].width = 12
        for col in range(5, 6 + self.num_choices):
            ws_stats.column_dimensions[get_column_letter(col)].width = 10

        wb.save(self.output_file)
        print(f"\n✓ 結果を保存しました: {self.output_file}")

    def print_summary(self, course_selection, schedule):
        """結果のサマリーを表示"""
        print("\n" + "=" * 70)
        print("配置結果サマリー")
        print("=" * 70)

        print("\n【時限・講座別人数】")
        for period in range(1, self.num_periods + 1):
            print(f"\n  {period}限:")
            for course in self.courses:
                count = sum(1 for s in self.students
                            if schedule[s['id']].get(period) == course)
                in_range = self.min_per_course <= count <= self.max_per_course
                status = "✓" if in_range else "!"
                print(f"    {course}: {count}名 {status}")

        print("\n【希望達成状況】")
        rank_counts = defaultdict(int)
        total_assignments = 0

        for student in self.students:
            selected = course_selection.get(student['id'], set())
            for course in selected:
                total_assignments += 1
                rank = self.get_preference_rank(student, course)
                if rank <= self.num_choices:
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
    print("        学生講座配置プログラム（ILP最適化版）")
    print("=" * 70)
    print()

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
            num_choices = int(input("講座数（希望順位の数）を入力してください（例: 6）: "))
            if num_choices > 0:
                break
            print("1以上の数値を入力してください。")
        except ValueError:
            print("数値を入力してください。")

    while True:
        try:
            num_periods = int(input(f"受講する講座数を入力してください（1〜{num_choices}）: "))
            if 1 <= num_periods <= num_choices:
                break
            print(f"1〜{num_choices}の数値を入力してください。")
        except ValueError:
            print("数値を入力してください。")

    # 平均人数を計算して表示
    avg_per_course = num_students / num_choices
    print(f"\n※ 1コマあたりの平均人数: {avg_per_course:.1f}名")

    while True:
        try:
            min_per_course = int(input("1コマあたりの最低人数を入力してください: "))
            if min_per_course >= 0:
                break
            print("0以上の数値を入力してください。")
        except ValueError:
            print("数値を入力してください。")

    while True:
        try:
            max_per_course = int(input("1コマあたりの最高人数を入力してください: "))
            if max_per_course >= min_per_course:
                break
            print(f"{min_per_course}以上の数値を入力してください。")
        except ValueError:
            print("数値を入力してください。")

    try:
        scheduler = StudentScheduler(num_students, num_periods, num_choices, min_per_course, max_per_course)

        print("\n" + "=" * 70)
        print("ステップ1: 入力ファイルの準備")
        print("=" * 70)
        scheduler.create_input_template()

        print("\n" + "=" * 70)
        print("ステップ2: データ入力")
        print("=" * 70)
        scheduler.open_excel_file(scheduler.input_file)
        scheduler.wait_for_file_close(scheduler.input_file)

        print("\n" + "=" * 70)
        print("ステップ3: データ処理")
        print("=" * 70)
        scheduler.load_data()

        print("\n" + "=" * 70)
        print("ステップ4: 最適化計算（ILP）")
        print("=" * 70)
        course_selection, schedule = scheduler.solve_with_ilp()

        scheduler.print_summary(course_selection, schedule)

        print("\n" + "=" * 70)
        print("ステップ5: 結果の保存")
        print("=" * 70)
        scheduler.save_results(course_selection, schedule)

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
