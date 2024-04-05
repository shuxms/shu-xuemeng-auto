import openpyxl
import tkinter as tk
from tkinter import filedialog, Radiobutton, StringVar, Entry, Label
import re

def read_excel_process_and_save_txt(input_file, title_level, user_text):
    workbook = openpyxl.load_workbook(input_file)
    sheet = workbook.active

    # 根据选择的标题级别确定第二层标题级别
    if title_level == "\\section":
        second_level = "\\subsection"
    elif title_level == "\\subsection":
        second_level = "\\subsubsection"

    content = ""
    groups = {}  # 用于存储相同行的分组

    for row in sheet.iter_rows(min_row=2, values_only=True):
        category = row[0]
        course_name = row[1]
        teacher_name = row[2]
        course_evaluation = row[3]

        key = (category, course_name, teacher_name)
        if key not in groups:
            groups[key] = []

        if course_evaluation:
            groups[key].append(course_evaluation)

    for key, evaluations in groups.items():
        category, course_name, teacher_name = key

        if category == '吐槽':
            content += f"{title_level}{{{course_name}}}\n"
            content += f"{second_level}{{{teacher_name}}}\n"
            content += "\\begin{itemize}\n"
            content += "  \\item \\textcolor{second}{\\textbf{吐槽}}\n"
        elif category == '推荐':
            content += f"{title_level}{{{course_name}}}\n"
            content += f"{second_level}{{{teacher_name}}}\n"
            content += "\\begin{itemize}\n"
            content += "  \\item \\textcolor{main}{\\textbf{推荐}}\n"

        if evaluations:
            content += "  \\begin{itemize}\n"
            for evaluation in evaluations:
                content += f"    \\item \\textcolor{{gray}}{{{user_text}:}} {evaluation}\n"
            content += "  \\end{itemize}\n"

        content += "\\end{itemize}\n\n"

    # Process the content
    processed_content = process_content(content)

    # Save the processed content to a file
    save_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text Files", "*.txt")])
    if save_path:
        with open(save_path, "w", encoding="utf-8") as file:
            file.write(processed_content)
        result_label.config(text=f"处理完成！保存在：{save_path}")
    else:
        result_label.config(text="未选择保存路径！")

def process_content(content):
    # 定义一个正则表达式模式，用于匹配需要添加句号的\item条目
    # 使用负向前瞻断言排除特定的\item条目
    # 使用负向后瞻断言确保末尾没有标点符号的\item条目才添加句号
    item_pattern = r'(\\item (?!\\textcolor{main}{\\textbf{推荐}}|\\textcolor{second}{\\textbf{吐槽}})[^\n]+?)(?<![\.!?。！？])\n'

    # 在匹配的\item条目末尾添加句号
    processed_content = re.sub(item_pattern, r'\1.\n', content)

    return processed_content

# GUI setup
root = tk.Tk()
root.title("Excel处理程序")

title_level = StringVar(value="\\section")

section_radio = Radiobutton(root, text="Section", variable=title_level, value="\\section")
section_radio.pack(anchor=tk.W)

subsection_radio = Radiobutton(root, text="Subsection", variable=title_level, value="\\subsection")
subsection_radio.pack(anchor=tk.W)

user_text_label = Label(root, text="输入文字:")
user_text_label.pack(anchor=tk.W)

user_text_entry = Entry(root)
user_text_entry.pack(anchor=tk.W)

process_button = tk.Button(root, text="选择Excel文件并处理", command=lambda: select_excel_file(title_level.get()))
process_button.pack(pady=20)

result_label = tk.Label(root, text="")
result_label.pack()

def select_excel_file(title_level):
    user_text = user_text_entry.get()  # 获取用户输入的文字
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        read_excel_process_and_save_txt(file_path, title_level, user_text)  # 传递用户输入的文字
    else:
        result_label.config(text="未选择文件！")

root.mainloop()
