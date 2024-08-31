import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import pandas as pd
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
    df = pd.read_excel(file_path)
    for index, row in df.iterrows():
        if '姓名' not in row:
          raise ValueError(f"输入表格第一列第一行需要以'姓名'二字开头，不要直接写人名")
        if '可值班日' not in row:
          raise ValueError(f"输入表格第二列第一行需要以'科值班日'二字开头，不要直接写日期")
        name = row['姓名']
        days = row['可值班日'].split('、')
        teachers[name] = [weekday_map[day] for day in days]
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
            raise ValueError(f"{date_str} 这一天没有老师能安排")

        # 找到当前排班次数最少的老师
        available_teachers.sort(key=lambda x: teacher_stats[x]['count'])
        teacher_name = available_teachers[0]

        note = "（调休）" if date_str in work_on_weekend else ""
        schedule.append((date_str, english_to_chinese_weekday[weekday] + note, teacher_name))
        teacher_stats[teacher_name]['days'].append(date_str)
        teacher_stats[teacher_name]['count'] += 1

    return schedule, teacher_stats

# 写入Excel文件
def write_schedule_to_excel(schedule, output_file):
    df = pd.DataFrame(schedule, columns=['日期', '星期', '人名'])
    df.to_excel(output_file, index=False)

# 写入教师统计信息到Excel文件
def write_teacher_stats_to_excel(teacher_stats, output_file):
    data = []
    for teacher, stats in teacher_stats.items():
        data.append((teacher, stats['count'], '、'.join(stats['days'])))
    df = pd.DataFrame(data, columns=['教师姓名', '排班天数', '排班日期'])
    df.to_excel(output_file, index=False)

# Tkinter UI
class ScheduleApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("排班生成器")
        self.geometry("600x400")

        self.input_file_label = tk.Label(self, text="教师信息文件:")
        self.input_file_label.pack()

        self.input_file_entry = tk.Entry(self, width=40)
        self.input_file_entry.pack()

        self.browse_button = tk.Button(self, text="浏览", command=self.browse_file)
        self.browse_button.pack()

        self.start_date_label = tk.Label(self, text="起始日期 (YYYY-MM-DD):")
        self.start_date_label.pack()

        self.start_date_entry = tk.Entry(self, width=20)
        self.start_date_entry.pack()

        self.end_date_label = tk.Label(self, text="结束日期 (YYYY-MM-DD):")
        self.end_date_label.pack()

        self.end_date_entry = tk.Entry(self, width=20)
        self.end_date_entry.pack()

        self.holidays_label = tk.Label(self, text="当前假日:")
        self.holidays_label.pack()

        self.holidays_listbox = tk.Listbox(self, width=40, height=10)
        self.holidays_listbox.pack()

        self.add_holiday_button = tk.Button(self, text="增加假日", command=self.add_holiday)
        self.add_holiday_button.pack()

        self.remove_holiday_button = tk.Button(self, text="删除假日", command=self.remove_holiday)
        self.remove_holiday_button.pack()

        self.work_on_weekend_label = tk.Label(self, text="当前调休日:")
        self.work_on_weekend_label.pack()

        self.work_on_weekend_listbox = tk.Listbox(self, width=40, height=10)
        self.work_on_weekend_listbox.pack()

        self.add_work_on_weekend_button = tk.Button(self, text="增加调休日", command=self.add_work_on_weekend)
        self.add_work_on_weekend_button.pack()

        self.remove_work_on_weekend_button = tk.Button(self, text="删除调休日", command=self.remove_work_on_weekend)
        self.remove_work_on_weekend_button.pack()

        self.generate_button = tk.Button(self, text="生成排班", command=self.generate_schedule)
        self.generate_button.pack()

        self.update_holidays_and_work_on_weekend_listboxes()

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        self.input_file_entry.delete(0, tk.END)
        self.input_file_entry.insert(0, file_path)

    def add_holiday(self):
        holiday = simpledialog.askstring("输入", "请输入假日日期 (YYYY-MM-DD):")
        if holiday:
            holidays.add(holiday)
            self.update_holidays_and_work_on_weekend_listboxes()

    def remove_holiday(self):
        selected = self.holidays_listbox.get(self.holidays_listbox.curselection())
        if selected:
            holidays.remove(selected)
            self.update_holidays_and_work_on_weekend_listboxes()

    def add_work_on_weekend(self):
        work_on_weekend_date = simpledialog.askstring("输入", "请输入调休日日期 (YYYY-MM-DD):")
        if work_on_weekend_date:
            work_on_weekend.add(work_on_weekend_date)
            self.update_holidays_and_work_on_weekend_listboxes()

    def remove_work_on_weekend(self):
        selected = self.work_on_weekend_listbox.get(self.work_on_weekend_listbox.curselection())
        if selected:
            work_on_weekend.remove(selected)
            self.update_holidays_and_work_on_weekend_listboxes()

    def update_holidays_and_work_on_weekend_listboxes(self):
        self.holidays_listbox.delete(0, tk.END)
        for holiday in sorted(holidays):
            self.holidays_listbox.insert(tk.END, holiday)

        self.work_on_weekend_listbox.delete(0, tk.END)
        for work_on_weekend_date in sorted(work_on_weekend):
            self.work_on_weekend_listbox.insert(tk.END, work_on_weekend_date)

    def generate_schedule(self):
        input_file = self.input_file_entry.get()
        start_date_str = self.start_date_entry.get()
        end_date_str = self.end_date_entry.get()

        if not input_file:
            messagebox.showerror("错误", "请选择教师信息文件")
            return

        if not start_date_str or not end_date_str:
            messagebox.showerror("错误", "请输入起始日期和结束日期")
            return

        try:
            start_date = datetime.strptime(start_date_str, '%Y-%m-%d')
            end_date = datetime.strptime(end_date_str, '%Y-%m-%d')
        except ValueError:
            messagebox.showerror("错误", "日期格式不正确，请使用 YYYY-MM-DD 格式")
            return

        try:
            teachers = read_teachers(input_file)
            schedule, teacher_stats = generate_schedule(teachers, start_date, end_date)
            schedule_output_file = "排班结果.xlsx"
            stats_output_file = "排班统计信息.xlsx"
            write_schedule_to_excel(schedule, schedule_output_file)
            write_teacher_stats_to_excel(teacher_stats, stats_output_file)
            messagebox.showinfo("成功", "排班生成成功，文件已保存在app所在目录，请查看！")
        except Exception as e:
            messagebox.showerror("错误", f"生成排班时出错: {e}")

if __name__ == "__main__":
    app = ScheduleApp()
    app.mainloop()