# student_scheduler.py
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import random
from collections import defaultdict
import copy
import os
import subprocess
import platform
import time

class StudentScheduler:
    def __init__(self, num_students, num_choices):
        self.num_students = num_students
        self.num_choices = num_choices
        self.students = []
        self.courses = set()
        self.input_file = "student_preferences.xlsx"
        self.output_file = "schedule_result.xlsx"
        
    def create_input_template(self):
        """入力用のテンプレートExcelファイルを作成"""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "アンケート入力"
        
        # スタイル定義
        header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        header_font = Font(bold=True, color='FFFFFF', size=12)
        border = Border(
            left=Side(style='thin', color='000000'),
            right=Side(style='thin', color='000000'),
            top=Side(style='thin', color='000000'),
            bottom=Side(style='thin', color='000000')
        )
        
        # ヘッダー行
        headers = ['生徒名'] + [f'第{i}希望' for i in range(1, self.num_choices + 1)]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(1, col, header)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = border
        
        # サンプルデータ（最初の3行）
        sample_data = [
            ['山田太郎', 'プログラミング', '美術', '音楽', '体育', '英会話', '料理'],
            ['佐藤花子', '美術', '音楽', '料理', 'プログラミング', '体育', '英会話'],
            ['鈴木一郎', '体育', 'プログラミング', '英会話', '美術', '料理', '音楽'],
        ]
        
        for row_idx, data in enumerate(sample_data[:min(3, self.num_students)], 2):
            for col_idx, value in enumerate(data[:len(headers)], 1):
                cell = ws.cell(row_idx, col_idx, value)
                cell.border = border
                cell.alignment = Alignment(horizontal='left', vertical='center')
                if col_idx == 1:  # 生徒名列
                    cell.fill = PatternFill(start_color='E7E6E6', end_color='E7E6E6', fill_type='solid')
        
        # 残りの行を空行で用意
        for row_idx in range(len(sample_data) + 2, self.num_students + 2):
            for col_idx in range(1, len(headers) + 1):
                cell = ws.cell(row_idx, col_idx, '')
                cell.border = border
                cell.alignment = Alignment(horizontal='left', vertical='center')
                if col_idx == 1:  # 生徒名列
                    cell.fill = PatternFill(start_color='E7E6E6', end_color='E7E6E6', fill_type='solid')
        
        # 列幅の調整
        ws.column_dimensions['A'].width = 15
        for col in range(2, len(headers) + 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 18
        
        # 行の高さ調整
        ws.row_dimensions[1].height = 25
        for row in range(2, self.num_students + 2):
            ws.row_dimensions[row].height = 20
        
        # 注意事項シート
        ws_info = wb.create_sheet("注意事項", 0)
        ws_info.column_dimensions['A'].width = 80
        
        info_texts = [
            "【使用方法】",
            "",
            "1. 「アンケート入力」シートに生徒の希望を入力してください",
            "2. 生徒名と希望講座をすべて入力したら、ファイルを保存して閉じてください",
            "3. プログラムが自動的に配置を計算し、結果を表示します",
            "",
            "【入力上の注意】",
            f"・生徒数: {self.num_students}名分入力してください",
            f"・希望数: 第1希望から第{self.num_choices}希望まで入力してください",
            "・講座名は正確に入力してください（表記ゆれがあると別講座として扱われます）",
            "・サンプルデータは上書きして使用してください",
            "",
            "【配置について】",
            "・4つの講座が選ばれ、1限から4限に配置されます",
            "・各時限の人数ができる限り均等になるよう調整されます",
            "・可能な限り上位の希望が尊重されます",
        ]
        
        for row, text in enumerate(info_texts, 1):
            cell = ws_info.cell(row, 1, text)
            if row == 1:
                cell.font = Font(bold=True, size=14, color='4472C4')
            elif text.startswith("【"):
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
                # ファイルが開けた = まだExcelが開いていない
                print(".", end="", flush=True)
                time.sleep(1)
            except (PermissionError, IOError):
                # ファイルがロックされた = Excelが開いた
                file_opened = True
                print("\n✓ Excelがファイルを開きました。編集してください。")
                break

        if not file_opened:
            # 30秒待ってもロックされなかった場合、ユーザーに確認
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
        
        # データ読み込み
        for row in ws.iter_rows(min_row=2, max_row=self.num_students + 1, values_only=True):
            if row[0] is None or str(row[0]).strip() == '':
                continue
                
            preferences = []
            for i in range(1, self.num_choices + 1):
                if row[i] is not None and str(row[i]).strip() != '':
                    preferences.append(str(row[i]).strip())
            
            if preferences:  # 希望が1つ以上ある場合のみ追加
                student = {
                    'name': str(row[0]).strip(),
                    'preferences': preferences
                }
                self.students.append(student)
                self.courses.update(preferences)
        
        wb.close()
        
        if len(self.students) == 0:
            raise ValueError("有効な生徒データが見つかりません")
        
        print(f"\n✓ 読み込み完了: {len(self.students)}名の生徒データ")
        print(f"✓ 講座数: {len(self.courses)}講座")
        
        # 講座一覧を表示
        print("\n【登録された講座一覧】")
        for i, course in enumerate(sorted(self.courses), 1):
            print(f"  {i}. {course}")
    
    def calculate_score(self, assignment):
        """配置の評価スコアを計算（小さいほど良い）"""
        total_score = 0
        for student in self.students:
            assigned_courses = assignment.get(student['name'], [])
            for course in assigned_courses:
                if course in student['preferences']:
                    rank = student['preferences'].index(course) + 1
                    total_score += rank
                else:
                    total_score += 100
        return total_score
    
    def calculate_balance_penalty(self, schedule):
        """時限間の人数バランスのペナルティを計算"""
        period_counts = [len(students) for students in schedule.values()]
        if not period_counts:
            return 0
        avg = sum(period_counts) / len(period_counts)
        variance = sum((count - avg) ** 2 for count in period_counts)
        return variance
    
    def greedy_assign(self):
        """貪欲法による初期配置"""
        schedule = {i: [] for i in range(1, 5)}
        assignment = {}
        
        # 講座の人気度を計算
        course_popularity = defaultdict(int)
        for student in self.students:
            if student['preferences']:
                course_popularity[student['preferences'][0]] += 1
        
        # 人気講座上位4つを選択
        sorted_courses = sorted(course_popularity.items(), 
                               key=lambda x: x[1], 
                               reverse=True)
        selected_courses = [course for course, _ in sorted_courses[:4]]
        
        # 講座を時限に割り当て
        course_to_period = {}
        for i, course in enumerate(selected_courses, 1):
            course_to_period[course] = i
        
        print(f"\n【選ばれた4講座】")
        for period, course in enumerate(selected_courses, 1):
            print(f"  {period}限: {course} (第1希望: {course_popularity[course]}名)")
        
        # 目標人数を計算
        target_per_period = len(self.students) / 4
        max_per_period = int(target_per_period + 5)
        
        # 生徒を配置
        unassigned_students = []
        
        for student in self.students:
            assigned = False
            for pref_course in student['preferences']:
                if pref_course in course_to_period:
                    period = course_to_period[pref_course]
                    if len(schedule[period]) < max_per_period:
                        schedule[period].append(student['name'])
                        assignment[student['name']] = [pref_course]
                        assigned = True
                        break
            
            if not assigned:
                unassigned_students.append(student)
        
        # 未配置の生徒を処理
        for student in unassigned_students:
            min_period = min(schedule.keys(), key=lambda p: len(schedule[p]))
            schedule[min_period].append(student['name'])
            
            # その時限の講座を割り当て
            period_course = [c for c, p in course_to_period.items() if p == min_period]
            if period_course:
                assignment[student['name']] = [period_course[0]]
            else:
                assignment[student['name']] = ['未配置']
        
        return schedule, assignment, course_to_period
    
    def improve_schedule(self, schedule, assignment, course_to_period, iterations=10000):
        """焼きなまし法で配置を改善"""
        best_schedule = copy.deepcopy(schedule)
        best_assignment = copy.deepcopy(assignment)
        best_score = self.calculate_score(best_assignment) + \
                     self.calculate_balance_penalty(best_schedule) * 10
        
        print(f"\n配置を最適化中", end="")
        
        for iteration in range(iterations):
            if iteration % 1000 == 0:
                print(".", end="", flush=True)
            
            periods = list(schedule.keys())
            p1, p2 = random.sample(periods, 2)
            
            if not schedule[p1] or not schedule[p2]:
                continue
            
            s1 = random.choice(schedule[p1])
            s2 = random.choice(schedule[p2])
            
            # 交換を試行
            new_schedule = copy.deepcopy(schedule)
            new_assignment = copy.deepcopy(assignment)
            
            new_schedule[p1].remove(s1)
            new_schedule[p1].append(s2)
            new_schedule[p2].remove(s2)
            new_schedule[p2].append(s1)
            
            # 講座の再割り当て
            for student_name, period in [(s1, p2), (s2, p1)]:
                student = next(s for s in self.students if s['name'] == student_name)
                period_courses = [c for c, p in course_to_period.items() if p == period]
                
                best_course = None
                best_rank = float('inf')
                for course in period_courses:
                    if course in student['preferences']:
                        rank = student['preferences'].index(course)
                        if rank < best_rank:
                            best_rank = rank
                            best_course = course
                
                if best_course:
                    new_assignment[student_name] = [best_course]
            
            # 新しいスコアを計算
            new_score = self.calculate_score(new_assignment) + \
                       self.calculate_balance_penalty(new_schedule) * 10
            
            # 改善されていれば採用
            if new_score < best_score:
                best_schedule = new_schedule
                best_assignment = new_assignment
                best_score = new_score
        
        print(" 完了!")
        return best_schedule, best_assignment
    
    def save_results(self, schedule, assignment, course_to_period):
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
        
        # シート1: サマリー
        ws_summary = wb.active
        ws_summary.title = "📊 サマリー"
        
        row = 1
        title = ws_summary.cell(row, 1, "配置結果サマリー")
        title.font = Font(bold=True, size=16, color='4472C4')
        row += 2
        
        # 時限別人数
        ws_summary.cell(row, 1, "時限別配置状況").font = Font(bold=True, size=12)
        row += 1
        
        headers = ['時限', '講座名', '人数', '割合']
        for col, header in enumerate(headers, 1):
            cell = ws_summary.cell(row, col, header)
            cell.fill = header_fill
            cell.font = header_font
            cell.border = border
            cell.alignment = Alignment(horizontal='center', vertical='center')
        row += 1
        
        total_students = len(self.students)
        for period in sorted(schedule.keys()):
            students = schedule[period]
            period_courses = [c for c, p in course_to_period.items() if p == period]
            course_name = period_courses[0] if period_courses else '未設定'
            count = len(students)
            percentage = count / total_students * 100
            
            ws_summary.cell(row, 1, f"{period}限").border = border
            ws_summary.cell(row, 2, course_name).border = border
            ws_summary.cell(row, 3, count).border = border
            ws_summary.cell(row, 4, f"{percentage:.1f}%").border = border
            
            # 人数に応じて色分け
            target = total_students / 4
            if abs(count - target) <= 2:
                fill = good_fill
            elif abs(count - target) <= 5:
                fill = warning_fill
            else:
                fill = bad_fill
            
            for col in range(1, 5):
                ws_summary.cell(row, col).fill = fill
                ws_summary.cell(row, col).alignment = Alignment(horizontal='center', vertical='center')
            
            row += 1
        
        row += 2
        
        # 希望達成状況
        ws_summary.cell(row, 1, "希望達成状況").font = Font(bold=True, size=12)
        row += 1
        
        headers = ['希望順位', '人数', '割合']
        for col, header in enumerate(headers, 1):
            cell = ws_summary.cell(row, col, header)
            cell.fill = header_fill
            cell.font = header_font
            cell.border = border
            cell.alignment = Alignment(horizontal='center', vertical='center')
        row += 1
        
        rank_counts = defaultdict(int)
        for student in self.students:
            if student['name'] in assignment:
                assigned_course = assignment[student['name']][0]
                if assigned_course in student['preferences']:
                    rank = student['preferences'].index(assigned_course) + 1
                    rank_counts[rank] += 1
                else:
                    rank_counts['希望外'] += 1
        
        for rank in range(1, self.num_choices + 1):
            count = rank_counts.get(rank, 0)
            percentage = count / total_students * 100
            
            ws_summary.cell(row, 1, f"第{rank}希望").border = border
            ws_summary.cell(row, 2, count).border = border
            ws_summary.cell(row, 3, f"{percentage:.1f}%").border = border
            
            # 順位に応じて色分け
            if rank <= 2:
                fill = good_fill
            elif rank <= 4:
                fill = warning_fill
            else:
                fill = bad_fill
            
            for col in range(1, 4):
                ws_summary.cell(row, col).fill = fill
                ws_summary.cell(row, col).alignment = Alignment(horizontal='center', vertical='center')
            
            row += 1
        
        # 希望外
        hope_outside = rank_counts.get('希望外', 0)
        if hope_outside > 0:
            percentage = hope_outside / total_students * 100
            ws_summary.cell(row, 1, "希望外").border = border
            ws_summary.cell(row, 2, hope_outside).border = border
            ws_summary.cell(row, 3, f"{percentage:.1f}%").border = border
            
            for col in range(1, 4):
                ws_summary.cell(row, col).fill = bad_fill
                ws_summary.cell(row, col).border = border
                ws_summary.cell(row, col).alignment = Alignment(horizontal='center', vertical='center')
        
        # 列幅調整
        ws_summary.column_dimensions['A'].width = 15
        ws_summary.column_dimensions['B'].width = 25
        ws_summary.column_dimensions['C'].width = 12
        ws_summary.column_dimensions['D'].width = 12
        
        # シート2: 時限別配置
        ws_period = wb.create_sheet("🕐 時限別配置")
        
        row = 1
        for period in sorted(schedule.keys()):
            students = schedule[period]
            period_courses = [c for c, p in course_to_period.items() if p == period]
            course_name = period_courses[0] if period_courses else '未設定'
            
            # 時限ヘッダー
            cell = ws_period.cell(row, 1, f"【{period}限】 {course_name}")
            cell.font = Font(bold=True, size=12, color='FFFFFF')
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='left', vertical='center')
            ws_period.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
            row += 1
            
            cell = ws_period.cell(row, 1, f"人数: {len(students)}名")
            cell.font = Font(bold=True)
            cell.fill = subheader_fill
            ws_period.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
            row += 1
            
            # 生徒リスト
            for i, student_name in enumerate(sorted(students), 1):
                ws_period.cell(row, 1, i).border = border
                ws_period.cell(row, 1).alignment = Alignment(horizontal='center')
                ws_period.cell(row, 2, student_name).border = border
                row += 1
            
            row += 1
        
        ws_period.column_dimensions['A'].width = 8
        ws_period.column_dimensions['B'].width = 20
        
        # シート3: 生徒別配置
        ws_student = wb.create_sheet("👥 生徒別配置")
        
        headers = ['No.', '生徒名', '配置講座', '希望順位'] + \
                  [f'第{i}希望' for i in range(1, self.num_choices + 1)]
        
        for col, header in enumerate(headers, 1):
            cell = ws_student.cell(1, col, header)
            cell.fill = header_fill
            cell.font = header_font
            cell.border = border
            cell.alignment = Alignment(horizontal='center', vertical='center')
        
        row = 2
        for idx, student in enumerate(sorted(self.students, key=lambda s: s['name']), 1):
            ws_student.cell(row, 1, idx).border = border
            ws_student.cell(row, 1).alignment = Alignment(horizontal='center')
            ws_student.cell(row, 2, student['name']).border = border
            
            if student['name'] in assignment:
                assigned_course = assignment[student['name']][0]
                ws_student.cell(row, 3, assigned_course).border = border
                
                # 希望順位を計算
                rank_cell = ws_student.cell(row, 4)
                rank_cell.border = border
                rank_cell.alignment = Alignment(horizontal='center')
                
                if assigned_course in student['preferences']:
                    rank = student['preferences'].index(assigned_course) + 1
                    rank_cell.value = f"第{rank}希望"
                    
                    if rank <= 2:
                        rank_cell.fill = good_fill
                    elif rank <= 4:
                        rank_cell.fill = warning_fill
                    else:
                        rank_cell.fill = bad_fill
                else:
                    rank_cell.value = "希望外"
                    rank_cell.fill = bad_fill
            
            # 希望を表示
            for i, pref in enumerate(student['preferences'], 5):
                ws_student.cell(row, i, pref).border = border
            
            row += 1
        
        # 列幅調整
        ws_student.column_dimensions['A'].width = 6
        ws_student.column_dimensions['B'].width = 15
        ws_student.column_dimensions['C'].width = 20
        ws_student.column_dimensions['D'].width = 12
        for col in range(5, 5 + self.num_choices):
            ws_student.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 18
        
        wb.save(self.output_file)
        print(f"\n✓ 結果を保存しました: {self.output_file}")

def main():
    print("=" * 70)
    print("　　　　　学生講座配置プログラム")
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
            num_choices = int(input("希望順位の数を入力してください（例: 6）: "))
            if num_choices > 0:
                break
            print("1以上の数値を入力してください。")
        except ValueError:
            print("数値を入力してください。")
    
    try:
        scheduler = StudentScheduler(num_students, num_choices)
        
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
        
        # 初期配置
        print("\n初期配置を作成中...")
        schedule, assignment, course_to_period = scheduler.greedy_assign()
        
        # 配置の改善
        schedule, assignment = scheduler.improve_schedule(
            schedule, assignment, course_to_period, iterations=10000
        )
        
        # 結果の表示
        print("\n" + "=" * 70)
        print("配置結果")
        print("=" * 70)
        
        print("\n【時限別人数】")
        for period in sorted(schedule.keys()):
            period_courses = [c for c, p in course_to_period.items() if p == period]
            course_name = period_courses[0] if period_courses else '未設定'
            print(f"  {period}限 ({course_name}): {len(schedule[period])}名")
        
        print("\n【希望達成状況】")
        rank_counts = defaultdict(int)
        for student in scheduler.students:
            if student['name'] in assignment:
                assigned_course = assignment[student['name']][0]
                if assigned_course in student['preferences']:
                    rank = student['preferences'].index(assigned_course) + 1
                    rank_counts[rank] += 1
                else:
                    rank_counts['希望外'] += 1
        
        for rank in range(1, num_choices + 1):
            count = rank_counts.get(rank, 0)
            percentage = count / len(scheduler.students) * 100
            bar = "■" * int(percentage / 5)
            print(f"  第{rank}希望: {count:3d}名 ({percentage:5.1f}%) {bar}")
        
        hope_outside = rank_counts.get('希望外', 0)
        if hope_outside > 0:
            percentage = hope_outside / len(scheduler.students) * 100
            bar = "■" * int(percentage / 5)
            print(f"  希望外 : {hope_outside:3d}名 ({percentage:5.1f}%) {bar}")
        
        # 結果を保存
        print("\n" + "=" * 70)
        print("ステップ4: 結果の保存")
        print("=" * 70)
        scheduler.save_results(schedule, assignment, course_to_period)
        
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