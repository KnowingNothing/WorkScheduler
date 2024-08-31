import tkinter as tk
from tkinter import filedialog, messagebox
import csv
from datetime import datetime, timedelta
from collections import defaultdict

# 节假日列表
holidays = {
    "2024-01-01",  # 元旦
    "2024-02-10", "2024-02-11", "2024-02-12", "2024-02-13", "2024-02-14", "2024-02-15", "2024-02-16", "2024-02-17",  # 春节
    "2024-04-04", "2024-04-05", "2024-04-06",  # 清明节
    "2024-05-01", "2024-05-02", "2024-05-03", "2024-05-04", "2024-05-05",  # 劳动节
    "2024-06-10",  # 端午节
    "2024-09-15", "2024-09-16", "2024-09-17",  # 中秋节
    "2024-10-01", "2024-10-02", "2024-10-03", "2024-10-04", "2024-10-05", "2024-10-06", "2024-10-07",  # 国庆节
    "2025-01-01"  # 2025年元旦
}

# 调休的周六日
work_on_weekend = {
    "2024-09-14", "2024-09-29", "2024-10-12", "2025-01-11"
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
    with open(file_path, 'r', encoding='utf-8') as file:
        for line in file:
            name, days = line.strip().split(' ', 1)
            teachers[name] = [weekday_map[day] for day in days.split('、')]
    return teachers

# 生成日期范围
def generate_dates(start_date, end_date):
    current_date = start_date
    while current_date <= end_date:
        yield current_date
        current_date += timedelta(days=1)

# 生成排班表
def generate_schedule(teachers, start_date, end_date):
    schedule = []
    teacher_stats = defaultdict(lambda: {'days': [], 'count': 0})
    teacher_list = list(teachers.keys())
    teacher_index = 0

    for date in generate_dates(start_date, end_date):
        date_str = date.strftime('%Y-%m-%d')
        weekday = date.strftime('%A')
        if date_str in holidays or (weekday in ['Saturday', 'Sunday'] and date_str not in work_on_weekend):
            continue

        # 找到可以值班的老师
        available_teachers = [teacher for teacher in teacher_list if weekday in teachers[teacher] or date_str in work_on_weekend]
        if not available_teachers:
            raise ValueError(f"No teacher available for date {date_str}")

        # 找到当前排班次数最少的老师
        available_teachers.sort(key=lambda x: teacher_stats[x]['count'])
        teacher_name = available_teachers[0]

        note = "（调休）" if date_str in work_on_weekend else ""
        schedule.append((date_str, english_to_chinese_weekday[weekday] + note, teacher_name))
        teacher_stats[teacher_name]['days'].append(date_str)
        teacher_stats[teacher_name]['count'] += 1

    return schedule, teacher_stats

# 写入CSV文件
def write_schedule_to_csv(schedule, output_file):
    with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
        csvwriter = csv.writer(csvfile)
        csvwriter.writerow(['日期', '星期', '人名'])
        csvwriter.writerows(schedule)

# 写入教师统计信息到CSV文件
def write_teacher_stats_to_csv(teacher_stats, output_file):
    with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
        csvwriter = csv.writer(csvfile)
        csvwriter.writerow(['教师姓名', '排班天数', '排班日期'])
        for teacher, stats in teacher_stats.items():
            csvwriter.writerow([teacher, stats['count'], '、'.join(stats['days'])])

# Tkinter UI
class ScheduleApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("排班生成器")
        self.geometry("400x200")

        self.input_file_label = tk.Label(self, text="教师信息文件:")
        self.input_file_label.pack()

        self.input_file_entry = tk.Entry(self, width=40)
        self.input_file_entry.pack()

        self.browse_button = tk.Button(self, text="浏览", command=self.browse_file)
        self.browse_button.pack()

        self.generate_button = tk.Button(self, text="生成排班", command=self.generate_schedule)
        self.generate_button.pack()

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
        self.input_file_entry.delete(0, tk.END)
        self.input_file_entry.insert(0, file_path)

    def generate_schedule(self):
        input_file = self.input_file_entry.get()
        if not input_file:
            messagebox.showerror("错误", "请选择教师信息文件")
            return

        try:
            teachers = read_teachers(input_file)
            start_date = datetime(2024, 9, 2)
            end_date = datetime(2025, 1, 11)
            schedule, teacher_stats = generate_schedule(teachers, start_date, end_date)
            schedule_output_file = "schedule_night.csv"
            stats_output_file = "teacher_stats_night.csv"
            write_schedule_to_csv(schedule, schedule_output_file)
            write_teacher_stats_to_csv(teacher_stats, stats_output_file)
            messagebox.showinfo("成功", "排班生成成功，文件已保存")
        except Exception as e:
            messagebox.showerror("错误", f"生成排班时出错: {e}")

if __name__ == "__main__":
    app = ScheduleApp()
    app.mainloop()