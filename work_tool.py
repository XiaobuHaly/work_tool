import os
from tkinter import *
from tkinter import messagebox
from openpyxl import load_workbook
from datetime import datetime

# 定义作业目录路径
directory = r"work_tool\data"

# 创建 Tkinter 应用
app = Tk()
app.title("作业查看工具")

# 设置窗口的最小大小
app.minsize(600, 300)

# 设置窗口的初始大小
app.geometry('600x650')

# 来源选项，0 表示文件的父目录，1 表示文件名
source_option = 1

# 判断一行是否在合并单元格中
def in_merged_cells(row, merged_cells):
    for m in merged_cells:
        if m.min_row <= row <= m.max_row:
            return True, f"{m.min_row}.{m.max_row}" # 返回是合并单元格及对应的日期范围
    return False, None # 是合并单元格及对应的日期范围

# 格式化日期范围
def format_date_range(date_range):
    min_date, max_date = date_range.split('.')
    if min_date == max_date:
        return min_date
    else:
        return date_range

# 判断一行是否在合并单元格中
def in_merged_cells(row, merged_cells):
    for m in merged_cells:
        if m.min_row <= row <= m.max_row:
            return True, m # 返回是合并单元格及对应的合并单元格对象
    return False, None # 是合并单元格及对应的合并单元格对象

# 读取特定日期的作业内容
def read_homework(date):
    homework_content = ""
    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith(".xlsx"):
                file_path = os.path.join(root, file)
                if source_option == 0:
                    source = os.path.basename(os.path.dirname(file_path))
                elif source_option == 1:
                    source = os.path.splitext(file)[0]
                wb = load_workbook(file_path)
                sheet = wb.active

                for i, row in enumerate(sheet.iter_rows(values_only=True), start=1):
                    try:
                        excel_date = row[0].strftime('%m.%d')
                    except AttributeError:
                        excel_date = str(row[0]) 

                    if excel_date != date:
                        continue

                    in_merge, date_range = in_merged_cells(i, sheet.merged_cells)
                    if in_merge:
                        homework_items = []
                        item_counter = 0
                        for j in range(date_range.min_row, date_range.max_row+1):
                            homework_item = [f"{idx+1 + item_counter}、{cell.value}" for idx, cell in enumerate(sheet[j][1:]) if cell.value]
                            homework_items.extend(homework_item)
                            item_counter += len(homework_item)
                    else:
                        homework_items = [f"{idx+1}、 {item}" for idx, item in enumerate(row[1:]) if item]

                    if homework_items:
                        homework_content += f"来源：{source}\n"
                        homework_content += '\n'.join(homework_items) + "\n\n"
    return homework_content


# 显示特定日期的作业内容
def show_homework():
    date = date_entry.get()
    if not date:
        messagebox.showerror("错误", "请输入日期！")
        return
    homework_text.delete(1.0, END)
    homework = read_homework(date)
    homework_text.insert(END, homework if homework else "当天无作业")

# 创建日期输入框和按钮
date_label = Label(app, text="请输入日期(格式：#m.#d，例如：1.29、2.3):")
date_label.grid(row=0, column=0, padx=10, pady=10, sticky=W+E)
date_entry = Entry(app)
date_entry.grid(row=0, column=1, padx=10, pady=10, sticky=W+E)
date_entry.insert(0, f"{datetime.today().month}.{datetime.today().day}")
show_button = Button(app, text="查看作业", command=show_homework)
show_button.grid(row=0, column=2, padx=10, pady=10, sticky=W+E)

# 创建作业显示文本框
homework_text = Text(app, width=50, height=10)
homework_text.grid(row=1, column=0, columnspan=3, padx=10, pady=10, sticky=W+E+N+S)

# 设置行和列的权重，使得它们可以随着窗口的大小变化而变化
app.grid_rowconfigure(1, weight=1)
app.grid_columnconfigure(1, weight=1)

# 运行
app.mainloop()