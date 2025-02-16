import sys
import pandas as pd
from datetime import datetime, timedelta
from collections import defaultdict
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, QFileDialog, QMessageBox, QListWidget, QTextEdit, QInputDialog
from PyQt6.QtCore import Qt

# 2025 年放假调休信息
holidays = {
    "2025-01-01",  # 元旦
    # 春节：1月28日（农历除夕、周二）至2月4日（农历正月初七、周二）放假调休，共8天。
    *[f"2025-01-{day:02d}" for day in range(28, 32)],
    *[f"2025-02-{day:02d}" for day in range(1, 5)],
    # 清明节：4月4日（周五）至6日（周日）放假，共3天。
    *[f"2025-04-{day:02d}" for day in range(4, 7)],
    # 劳动节：5月1日（周四）至5日（周一）放假调休，共5天。
    *[f"2025-05-{day:02d}" for day in range(1, 6)],
    # 端午节：5月31日（周六）至6月2日（周一）放假，共3天。
    *[f"2025-05-{day:02d}" for day in range(31, 32)],
    *[f"2025-06-{day:02d}" for day in range(1, 3)],
    # 国庆节、中秋节：10月1日（周三）至8日（周三）放假调休，共8天。
    *[f"2025-10-{day:02d}" for day in range(1, 9)]
}

work_on_weekend = {
    "2025-01-26",  # 春节调休上班
    "2025-02-08",
    "2025-04-27",  # 劳动节调休上班
    "2025-09-28",  # 国庆节、中秋节调休上班
    "2025-10-11",
    "2025-07-06",  # 学校特殊安排
    "2025-07-13"   # 学校特殊安排
}

# 汉字星期到英文星期的映射
weekday_map = {
    '一': 'Monday',
    '二': 'Tuesday',
    '三': 'Wednesday',
    '四': 'Thursday',
    '五': 'Friday',
    '六': 'Saturday',
    '日': 'Sunday'
}

# 英文星期到汉字星期的映射
english_to_chinese_weekday = {
    'Monday': '一',
    'Tuesday': '二',
    'Wednesday': '三',
    'Thursday': '四',
    'Friday': '五',
    'Saturday': '六',
    'Sunday': '日'
}

# 读取教师信息
def read_teachers(file_path):
    teachers = {}
    df = pd.read_excel(file_path)
    for index, row in df.iterrows():
        if '姓名' not in row:
            raise ValueError(f"输入表格第一列第一行需要以'姓名'二字开头，不要直接写人名")
        if '可值班日' not in row:
            raise ValueError(f"输入表格第二列第一行需要以'可值班日'二字开头，不要直接写日期")
        name = row['姓名'].strip()
        days_str = row['可值班日'].strip()
        # 处理输入错误，将英文顿号替换为中文顿号并去除空格
        days_str = days_str.replace(',', '、').replace(' ', '')
        days = days_str.split('、') if days_str else []
        teachers[name] = [weekday_map[day] for day in days if day in weekday_map]
    return teachers

# 生成日期范围
def generate_dates(start_date, end_date):
    current_date = start_date
    while current_date <= end_date:
        yield current_date
        current_date += timedelta(days=1)

# 生成排班表（区分上下午）
def generate_schedule(morning_teachers, afternoon_teachers, start_date, end_date):
    morning_schedule = []
    afternoon_schedule = []
    teacher_stats = defaultdict(lambda: {'days': [], 'count': 0})

    for date in generate_dates(start_date, end_date):
        date_str = date.strftime('%Y-%m-%d')
        weekday = date.strftime('%A')
        if date_str in holidays or (weekday in ['Saturday', 'Sunday'] and date_str not in work_on_weekend):
            continue

        # 上午排班
        available_morning_teachers = [teacher for teacher in morning_teachers if weekday in morning_teachers[teacher] or date_str in work_on_weekend]
        if not available_morning_teachers:
            raise ValueError(f"{date_str} 上午没有老师能安排")
        available_morning_teachers.sort(key=lambda x: teacher_stats[x]['count'])
        morning_teacher_name = available_morning_teachers[0]
        morning_note = "（调休）" if date_str in work_on_weekend else ""
        morning_schedule.append((date_str, english_to_chinese_weekday[weekday] + morning_note, morning_teacher_name))
        teacher_stats[morning_teacher_name]['days'].append(date_str)
        teacher_stats[morning_teacher_name]['count'] += 1

        # 下午排班
        available_afternoon_teachers = [teacher for teacher in afternoon_teachers if weekday in afternoon_teachers[teacher] or date_str in work_on_weekend]
        if not available_afternoon_teachers:
            raise ValueError(f"{date_str} 下午没有老师能安排")
        available_afternoon_teachers.sort(key=lambda x: teacher_stats[x]['count'])
        afternoon_teacher_name = available_afternoon_teachers[0]
        afternoon_note = "（调休）" if date_str in work_on_weekend else ""
        afternoon_schedule.append((date_str, english_to_chinese_weekday[weekday] + afternoon_note, afternoon_teacher_name))
        teacher_stats[afternoon_teacher_name]['days'].append(date_str)
        teacher_stats[afternoon_teacher_name]['count'] += 1

    return morning_schedule, afternoon_schedule, teacher_stats

# 写入 Excel 文件
def write_schedule_to_excel(schedule, output_file):
    df = pd.DataFrame(schedule, columns=['日期', '星期', '人名'])
    df.to_excel(output_file, index=False)

# 写入教师统计信息到 Excel 文件
def write_teacher_stats_to_excel(teacher_stats, output_file):
    data = []
    for teacher, stats in teacher_stats.items():
        data.append((teacher, stats['count'], '、'.join(stats['days'])))
    df = pd.DataFrame(data, columns=['教师姓名', '排班天数', '排班日期'])
    df.to_excel(output_file, index=False)

class ScheduleApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        # 上午教师信息
        self.morning_input_file_label = QLabel("上午教师信息文件:")
        layout.addWidget(self.morning_input_file_label)
        self.morning_input_file_entry = QLineEdit()
        layout.addWidget(self.morning_input_file_entry)
        self.morning_browse_button = QPushButton("浏览")
        self.morning_browse_button.clicked.connect(self.browse_morning_file)
        layout.addWidget(self.morning_browse_button)

        # 下午教师信息
        self.afternoon_input_file_label = QLabel("下午教师信息文件:")
        layout.addWidget(self.afternoon_input_file_label)
        self.afternoon_input_file_entry = QLineEdit()
        layout.addWidget(self.afternoon_input_file_entry)
        self.afternoon_browse_button = QPushButton("浏览")
        self.afternoon_browse_button.clicked.connect(self.browse_afternoon_file)
        layout.addWidget(self.afternoon_browse_button)

        self.start_date_label = QLabel("起始日期 (YYYY-MM-DD):")
        layout.addWidget(self.start_date_label)
        self.start_date_entry = QLineEdit()
        self.start_date_entry.setText("2025-02-17")  # 默认起始日期
        layout.addWidget(self.start_date_entry)

        self.end_date_label = QLabel("结束日期 (YYYY-MM-DD):")
        layout.addWidget(self.end_date_label)
        self.end_date_entry = QLineEdit()
        self.end_date_entry.setText("2025-07-13")    # 默认结束日期
        layout.addWidget(self.end_date_entry)

        self.holidays_label = QLabel("当前假日:")
        layout.addWidget(self.holidays_label)
        self.holidays_listbox = QListWidget()
        layout.addWidget(self.holidays_listbox)
        self.add_holiday_button = QPushButton("增加假日")
        self.add_holiday_button.clicked.connect(self.add_holiday)
        layout.addWidget(self.add_holiday_button)
        self.remove_holiday_button = QPushButton("删除假日")
        self.remove_holiday_button.clicked.connect(self.remove_holiday)
        layout.addWidget(self.remove_holiday_button)

        self.work_on_weekend_label = QLabel("当前调休日:")
        layout.addWidget(self.work_on_weekend_label)
        self.work_on_weekend_listbox = QListWidget()
        layout.addWidget(self.work_on_weekend_listbox)
        self.add_work_on_weekend_button = QPushButton("增加调休日")
        self.add_work_on_weekend_button.clicked.connect(self.add_work_on_weekend)
        layout.addWidget(self.add_work_on_weekend_button)
        self.remove_work_on_weekend_button = QPushButton("删除调休日")
        self.remove_work_on_weekend_button.clicked.connect(self.remove_work_on_weekend)
        layout.addWidget(self.remove_work_on_weekend_button)

        self.generate_button = QPushButton("生成排班")
        self.generate_button.clicked.connect(self.generate_schedule)
        layout.addWidget(self.generate_button)

        self.schedule_result_text = QTextEdit()
        layout.addWidget(self.schedule_result_text)

        self.setLayout(layout)
        self.setWindowTitle('排班生成器')
        self.setGeometry(300, 300, 600, 700)

        self.update_holidays_and_work_on_weekend_listboxes()

    def browse_morning_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "选择上午教师信息文件", "", "Excel files (*.xlsx)")
        if file_path:
            self.morning_input_file_entry.setText(file_path)

    def browse_afternoon_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "选择下午教师信息文件", "", "Excel files (*.xlsx)")
        if file_path:
            self.afternoon_input_file_entry.setText(file_path)

    def add_holiday(self):
        holiday, ok = QInputDialog.getText(self, "输入", "请输入假日日期 (YYYY-MM-DD):")
        if ok and holiday:
            holidays.add(holiday)
            self.update_holidays_and_work_on_weekend_listboxes()

    def remove_holiday(self):
        selected_items = self.holidays_listbox.selectedItems()
        if selected_items:
            holiday = selected_items[0].text()
            holidays.remove(holiday)
            self.update_holidays_and_work_on_weekend_listboxes()

    def add_work_on_weekend(self):
        work_on_weekend_date, ok = QInputDialog.getText(self, "输入", "请输入调休日日期 (YYYY-MM-DD):")
        if ok and work_on_weekend_date:
            work_on_weekend.add(work_on_weekend_date)
            self.update_holidays_and_work_on_weekend_listboxes()

    def remove_work_on_weekend(self):
        selected_items = self.work_on_weekend_listbox.selectedItems()
        if selected_items:
            work_on_weekend_date = selected_items[0].text()
            work_on_weekend.remove(work_on_weekend_date)
            self.update_holidays_and_work_on_weekend_listboxes()

    def update_holidays_and_work_on_weekend_listboxes(self):
        self.holidays_listbox.clear()
        for holiday in sorted(holidays):
            self.holidays_listbox.addItem(holiday)

        self.work_on_weekend_listbox.clear()
        for work_on_weekend_date in sorted(work_on_weekend):
            self.work_on_weekend_listbox.addItem(work_on_weekend_date)

    def generate_schedule(self):
        morning_input_file = self.morning_input_file_entry.text()
        afternoon_input_file = self.afternoon_input_file_entry.text()
        start_date_str = self.start_date_entry.text()
        end_date_str = self.end_date_entry.text()

        if not morning_input_file or not afternoon_input_file:
            QMessageBox.critical(self, "错误", "请选择上午和下午教师信息文件")
            return

        if not start_date_str or not end_date_str:
            QMessageBox.critical(self, "错误", "请输入起始日期和结束日期")
            return

        try:
            start_date = datetime.strptime(start_date_str, '%Y-%m-%d')
            end_date = datetime.strptime(end_date_str, '%Y-%m-%d')
        except ValueError:
            QMessageBox.critical(self, "错误", "日期格式不正确，请使用 YYYY-MM-DD 格式")
            return

        try:
            morning_teachers = read_teachers(morning_input_file)
            afternoon_teachers = read_teachers(afternoon_input_file)
            morning_schedule, afternoon_schedule, teacher_stats = generate_schedule(
                morning_teachers, afternoon_teachers, start_date, end_date)

            morning_schedule_output_file = "上午排班结果.xlsx"
            afternoon_schedule_output_file = "下午排班结果.xlsx"
            stats_output_file = "排班统计信息.xlsx"

            write_schedule_to_excel(morning_schedule, morning_schedule_output_file)
            write_schedule_to_excel(afternoon_schedule, afternoon_schedule_output_file)
            write_teacher_stats_to_excel(teacher_stats, stats_output_file)

            QMessageBox.information(self, "成功", "排班生成成功，文件已保存在 app 所在目录，请查看！")

            # 显示排班结果到 UI
            self.schedule_result_text.clear()
            self.schedule_result_text.append("上午排班结果：")
            for date_str, weekday, teacher_name in morning_schedule:
                self.schedule_result_text.append(f"{date_str} {weekday} {teacher_name}")

            self.schedule_result_text.append("\n下午排班结果：")
            for date_str, weekday, teacher_name in afternoon_schedule:
                self.schedule_result_text.append(f"{date_str} {weekday} {teacher_name}")

        except Exception as e:
            QMessageBox.critical(self, "错误", f"生成排班时出错: {str(e)}")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ScheduleApp()
    ex.show()
    sys.exit(app.exec())