# ------------------- 版本号 -------------------
APP_VERSION = "v1.6.0"

import sys
import os
import re
import pandas as pd
from PyQt5.QtGui import QColor, QTextCursor, QTextCharFormat, QPixmap, QPainter, QFont, QTextDocument
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QPushButton, QTableWidget,
    QTableWidgetItem, QHeaderView, QHBoxLayout, QComboBox, QTextEdit,
    QMessageBox, QLineEdit
)
from PyQt5.QtCore import Qt, QDate, QRect
import os
import json

# 在全局（所有函数外）定义
# 获取当前程序运行目录（兼容打包 exe）
if getattr(sys, 'frozen', False):
    base_path = os.path.dirname(sys.executable)
else:
    base_path = os.getcwd()
json_path = os.path.join(base_path, "replace_rules.json")

# 如果文件存在则读取，否则自动创建默认模板
if os.path.exists(json_path):
    with open(json_path, 'r', encoding='utf-8') as f:
        replace_rules_dict = json.load(f)
else:
    # 定义默认规则
    default_rules = {
        "备1": "备1-可休",
        "十四": "十四 8.00~16.45",
        "夜": "夜 21.30~7.30",
        "晚": "晚 15.00~22.30",
        "早": "早 7.30~15.00",
        "中": "中 15.00~21.30",
        "四": "四 8.00~16.45",
        "五": "五 8.00~12.00",
        "二": "二 8.00~16.45",
        "三": "三 8.00~16.45",
        "九": "九 8.00~16.45",
        "十": "十 8.00~16.45",
        "1": "1 8.00~16.45",
        "2": "2 7.00-15.00",
        "3": "3 7.00-15.00",
        "4": "4 7.30~17.15",
        "5": "5 8.00~17.15",
        "6": "6 7.30~16.45",
        "7": "7 8.30~17.15",
        "9": "9 7.30~16.55",
        "15": "15 8.00~12.00",
        "西": "西 8.00~12.00",
        "备": "备 7.30~16.30",
        "帮": "帮 8.30~16.15",
        "休": "休息",
        "工": "工休"
    }

    # 写入文件
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(default_rules, f, ensure_ascii=False, indent=4)

    replace_rules_dict = default_rules  # 同时加载到内存
    print(f"未检测到 replace_rules.json 已自动创建默认配置文件 -> {json_path}")
    
# ------------------- Excel -> TXT -------------------
replace_rules = [
]
from lunarcalendar import Converter, Solar, Lunar

# ------------------ 配色方案 ------------------
COLOR_SCHEMES = {
    "方案1": {  # 柔和风格
        "COLOR_REST": QColor("#A6EBC3"),       
        "COLOR_NIGHT": QColor("#DDE6F2"),      
        "FONT_COLOR_HOLIDAY": QColor("#FF6347")
    },
    "方案2": {  # 活力 / 多巴胺风格
        "COLOR_REST": QColor("#27AF27"),       # 明亮绿（LimeGreen）
        "COLOR_NIGHT": QColor("#629DD8"),      # 明亮蓝（DodgerBlue）
        "FONT_COLOR_HOLIDAY": QColor("#FF4500") # 鲜橙红（OrangeRed）
    },
    "方案3": {  # 柔暖风格
        "COLOR_REST": QColor("#FFECB3"),       
        "COLOR_NIGHT": QColor("#FFCDD2"),      
        "FONT_COLOR_HOLIDAY": QColor("#D32F2F")
    },
    "方案4": {  # 沉稳风格
        "COLOR_REST": QColor("#80CBC4"),       
        "COLOR_NIGHT": QColor("#546E7A"),      
        "FONT_COLOR_HOLIDAY": QColor("#FF7043")
    }
}

# 月份 -> 配色方案映射
MONTH_COLOR_MAP = {
    1: "方案1",
    2: "方案2",
    3: "方案3",
    4: "方案4",
    5: "方案1",
    6: "方案2",
    7: "方案3",
    8: "方案4",
    9: "方案1",
    10: "方案2",
    11: "方案3",
    12: "方案4"
}

LINE_COLOR = "#E0E0E0"
TITLE_BG = "#F7F7F7"
FONT_FAMILY = "Microsoft YaHei"  # 或者 "PingFang" / "SimHei" / "Source Han Sans"
DATE_FONT = QFont(FONT_FAMILY, 10, QFont.Bold)
LUNAR_FONT = QFont(FONT_FAMILY, 9)
BODY_FONT = QFont(FONT_FAMILY, 10)
NOTE_FONT = QFont(FONT_FAMILY, 9)

CELL_PADDING = 8

# 农历数字转中文
MONTH_NAMES = ["正月", "二月", "三月", "四月", "五月", "六月",
               "七月", "八月", "九月", "十月", "冬月", "腊月"]
DAY_NAMES = ["初一", "初二", "初三", "初四", "初五", "初六", "初七", "初八", "初九", "初十",
             "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十",
             "廿一", "廿二", "廿三", "廿四", "廿五", "廿六", "廿七", "廿八", "廿九", "三十"]

def get_lunar_label(year, month, day):
    solar = Solar(year, month, day)
    lunar = Converter.Solar2Lunar(solar)

    # 如果初一 → 显示月份
    if lunar.day == 1:
        return MONTH_NAMES[lunar.month - 1]
    else:
        return DAY_NAMES[lunar.day - 1]

HOLIDAY_DICT = {
    # 2025年（闰六月）
    "2025-01-01": "元旦",
    "2025-02-08": "小年",
    "2025-02-09": "除夕",
    "2025-02-10": "春节",
    "2025-02-14": "情人节",
    "2025-02-24": "元宵节",
    "2025-04-05": "清明节",
    "2025-05-01": "劳动节",
    "2025-06-22": "端午节",
    "2025-08-22": "七夕节",
    "2025-10-01": "国庆节",
    "2025-10-06": "中秋节",
    "2025-10-31": "万圣节",
    "2025-12-22": "冬至",
    "2025-12-25": "圣诞节",
    # 2026年
    "2026-01-01": "元旦",
    "2026-01-28": "小年",
    "2026-01-29": "除夕",
    "2026-01-30": "春节",
    "2026-02-14": "情人节",
    "2026-02-13": "元宵节",
    "2026-04-05": "清明节",
    "2026-05-01": "劳动节",
    "2026-06-14": "端午节",
    "2026-08-11": "七夕节",
    "2026-09-10": "中秋节",
    "2026-10-01": "国庆节",
    "2026-10-31": "万圣节",
    "2026-12-21": "冬至",
    "2026-12-25": "圣诞节",
    # 2027年
    "2027-01-01": "元旦",
    "2027-02-16": "小年",
    "2027-02-17": "除夕",
    "2027-02-18": "春节",
    "2027-02-14": "情人节",
    "2027-03-04": "元宵节",
    "2027-04-05": "清明节",
    "2027-05-01": "劳动节",
    "2027-06-05": "端午节",
    "2027-07-31": "七夕节",
    "2027-09-30": "中秋节",
    "2027-10-01": "国庆节",
    "2027-10-31": "万圣节",
    "2027-12-21": "冬至",
    "2027-12-25": "圣诞节",
    # 2028年
    "2028-01-01": "元旦",
    "2028-02-04": "小年",
    "2028-02-05": "除夕",
    "2028-02-06": "春节",
    "2028-02-14": "情人节",
    "2028-02-22": "元宵节",
    "2028-04-05": "清明节",
    "2028-05-01": "劳动节",
    "2028-06-22": "端午节",
    "2028-08-19": "七夕节",
    "2028-09-18": "中秋节",
    "2028-10-01": "国庆节",
    "2028-10-31": "万圣节",
    "2028-12-21": "冬至",
    "2028-12-25": "圣诞节"
}

def generate_txt(files, month):
    if not files:
        return None
    current_directory = os.getcwd()
    output_dir = os.path.join(current_directory, "output_txt")  # TXT 子文件夹
    os.makedirs(output_dir, exist_ok=True)
    txt_file_paths = []

    def replace_column_in_text(input_file, replace_rules, column_number):
        with open(input_file, 'r', encoding='utf-8-sig') as file:
            lines = file.readlines()
        modified_lines = []
        for line in lines:
            if line.startswith("日期") or line.startswith("星期") or line.startswith("章学亭"):
                continue
            columns = line.strip().split('\t')
            if len(columns) >= column_number:
                # 遍历字典
                for target_word, replace_word in sorted(replace_rules_dict.items(), key=lambda x: -len(x[0])):
                    columns[column_number - 1] = re.sub(
                        rf'(?<!\S){re.escape(target_word)}(?!\S)', replace_word, columns[column_number - 1]
                    )
                # 清理多余空格
                columns[column_number - 1] = re.sub(r'\s+', ' ', columns[column_number - 1]).strip()
            modified_lines.append('\t'.join(columns) + '\n')
        return modified_lines

    for file_path in files:
        try:
            df = pd.read_excel(file_path, header=1)
        except:
            continue
        df.columns = df.columns.str.strip()
        if '星期' in df.columns:
            df['星期'] = df['星期'].apply(lambda x: '    ' + x if isinstance(x, str) else x)
        name = window.name_input.text() if hasattr(window, 'name_input') else None
        columns_to_keep = ['日期', '星期']
        if name:
            # 去除输入名字中的空格
            name_no_space = name.replace(' ', '')
            # 查找匹配的列名（忽略空格）
            for col in df.columns:
                if col.replace(' ', '') == name_no_space:
                    columns_to_keep.append(col)
                    break
        columns_to_keep = [col for col in columns_to_keep if col in df.columns]
        df = df[columns_to_keep]

        temp_txt = os.path.join(output_dir, os.path.basename(file_path).split('.')[0] + '_output.txt')
        df.to_csv(temp_txt, index=False, encoding='utf-8-sig', sep='\t')
        txt_file_paths.append(temp_txt)

    all_modified_lines = []
    for txt_file in txt_file_paths:
        all_modified_lines.extend(replace_column_in_text(txt_file, replace_rules, column_number=3))

    name = window.name_input.text() if hasattr(window, 'name_input') else "排班表"
    final_output_path = os.path.join(output_dir, f'{name}的{month}月排班表.txt')
    with open(final_output_path, 'w', encoding='utf-8-sig') as f:
        f.writelines(all_modified_lines)

    return final_output_path

# ------------------- TXT -> 排班字典 + 备注 -------------------
def process_txt_files(file_paths, year=None, month=None):
    schedule_dict = {}
    remarks = []
    for file_path in file_paths:
        with open(file_path, 'r', encoding='utf-8-sig') as f:
            lines = f.readlines()
        for line in lines:
            line = line.strip()
            if not line:
                continue
            parts = line.split('\t')
            if parts[0].startswith("日期") or parts[0].startswith("星期") or parts[0].startswith("章学亭"):
                continue
            if len(parts) < 3:
                remarks.append(line)
                continue
            date_raw = parts[0].strip()
            content = parts[2].strip()
            try:
                date_str = pd.to_datetime(date_raw, errors='coerce').strftime("%Y-%m-%d")
            except:
                date_str = None
            if (not date_str or date_str=="NaT") and date_raw.isdigit() and year is not None and month is not None:
                date_str = QDate(year, month, int(date_raw)).toString("yyyy-MM-dd")
            if not date_str:
                remarks.append(line)
                continue
            if date_str in schedule_dict:
                schedule_dict[date_str] += " | " + content
            else:
                schedule_dict[date_str] = content
    return schedule_dict, remarks

# ------------------- PyQt5 界面 -------------------
class ScheduleApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"附二院排班表生成器 ({APP_VERSION})")
        self.resize(900, 650)
        layout = QVBoxLayout()

        self.drag_label = QLabel("把 Excel 文件拖到这里，生成排班表")
        self.drag_label.setAlignment(Qt.AlignCenter)
        self.drag_label.setStyleSheet("border: 2px dashed #aaa; font-size: 18px; padding: 10px;")
        layout.addWidget(self.drag_label)

        # 年月和姓名选择
        input_layout = QHBoxLayout()
        
        # 姓名输入
        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("请输入姓名")
        self.name_input.setText("章学亭")  # 默认值
        
        # 年月选择
        self.year_combo = QComboBox()
        current_year = QDate.currentDate().year()
        for y in range(current_year-5, current_year+6):
            self.year_combo.addItem(str(y), y)
        self.year_combo.setCurrentText(str(current_year))

        self.month_combo = QComboBox()
        for m in range(1, 13):
            self.month_combo.addItem(f"{m}月", m)
        self.month_combo.setCurrentIndex(QDate.currentDate().month()-1)

        input_layout.addWidget(QLabel("姓名:"))
        input_layout.addWidget(self.name_input)
        input_layout.addWidget(QLabel("年份:"))
        input_layout.addWidget(self.year_combo)
        input_layout.addWidget(QLabel("月份:"))
        input_layout.addWidget(self.month_combo)
        layout.addLayout(input_layout)

        # 按钮
        btn_layout = QHBoxLayout()
        self.generate_btn = QPushButton("生成排班表")
        btn_layout.addWidget(self.generate_btn)
        self.save_btn = QPushButton("保存为 PNG 图片")
        btn_layout.addWidget(self.save_btn)
        layout.addLayout(btn_layout)

        # 表格
        self.table = QTableWidget()
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.table)

        # 备注
        self.remark_box = QTextEdit()
        self.remark_box.setReadOnly(True)
        self.remark_box.setFixedHeight(80)
        self.remark_box.setStyleSheet("color: gray; font-size: 12px; border-top: 1px solid #aaa;")
        layout.addWidget(self.remark_box)

        self.setLayout(layout)
        self.setAcceptDrops(True)

        self.files = []
        self.current_schedule_dict = {}
        self.current_remarks = []
        self.generate_btn.clicked.connect(self.on_generate)
        self.save_btn.clicked.connect(self.on_save_png)

    # 拖拽
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        self.files = []
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path.endswith(".xlsx"):
                self.files.append(path)
        if not self.files:
            QMessageBox.warning(self, "提示", "未检测到 Excel 文件，请重新拖入")
        else:
            self.drag_label.setText("\n".join([os.path.basename(f) for f in self.files]))

    # 生成表格
    def on_generate(self):
        if not self.files:
            QMessageBox.warning(self, "提示", "请先拖入 Excel 文件")
            return
        month = self.month_combo.currentData()
        txt_path = generate_txt(self.files, month)
        if not txt_path:
            QMessageBox.warning(self, "提示", "TXT 文件生成失败")
            return
        year = self.year_combo.currentData()
        schedule_dict, remarks = process_txt_files([txt_path], year, month)
        self.current_schedule_dict = schedule_dict
        self.current_remarks = remarks
        self.remark_box.setText("\n".join(remarks) if remarks else "")
        self.generate_schedule_table(schedule_dict, year, month)

    # 生成表格内容
    def generate_schedule_table(self, schedule_dict, year, month):
        first_day = QDate(year, month, 1)
        days_in_month = first_day.daysInMonth()
        total_rows = ((days_in_month + first_day.dayOfWeek() - 2)//7) + 1
        self.table.clear()
        self.table.setRowCount(total_rows)
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels(['周一','周二','周三','周四','周五','周六','周日'])

        start_col = (first_day.dayOfWeek()-1) % 7
        day = 1

        # ----------------- 获取当月配色方案 -----------------
        scheme_name = MONTH_COLOR_MAP.get(month, "方案1")
        scheme = COLOR_SCHEMES[scheme_name]
        COLOR_REST = scheme["COLOR_REST"]
        COLOR_NIGHT = scheme["COLOR_NIGHT"]
        FONT_COLOR_HOLIDAY = scheme["FONT_COLOR_HOLIDAY"]

        # 更新 color_map
        color_map = {"休息": COLOR_REST,
                     "工休": COLOR_REST, 
                     "夜": COLOR_NIGHT, 
                     "21.30~7.30": COLOR_NIGHT
                     }

        for row in range(total_rows):
            for col in range(7):
                if row == 0 and col < start_col:
                    self.table.setCellWidget(row, col, QLabel(""))
                    continue
                if day > days_in_month:
                    self.table.setCellWidget(row, col, QLabel(""))
                    continue

                date_str = QDate(year, month, day).toString("yyyy-MM-dd")
                schedule_text = schedule_dict.get(date_str, '')
                holiday_name = HOLIDAY_DICT.get(date_str, '')

                # 创建标签用于显示富文本
                label = QLabel()
                label.setWordWrap(True)  # 允许文本换行
                label.setMargin(2)  # 设置边距

                # 构建HTML文本
                html_text = []
                # 第一行：日期数字（和节假日）
                first_line = f"<span style='color: black;'>{day}</span>"
                if holiday_name:
                    first_line += f" <span style='color: #FF6347;'>{holiday_name}</span>"
                html_text.append(first_line)
                
                # 最后一行：排班内容
                if schedule_text:
                    # 将班次和时间分开，假设格式为"班次 时间"
                    parts = schedule_text.split(' ', 1)  # 只分割第一个空格
                    if len(parts) > 1:
                        班次, 时间 = parts
                        html_text.append(f"<span style='color: black;'>{班次}</span> <span style='color: #808080;'>{时间}</span>")
                    else:
                        # 如果没有时间部分（比如"休息"），就全部用黑色显示
                        html_text.append(f"<span style='color: black;'>{schedule_text}</span>")
                
                # 设置HTML文本
                label.setText("<br>".join(html_text))
                
                # 设置单元格部件
                self.table.setCellWidget(row, col, label)

                # 设置背景色（排班相关）
                if schedule_text:
                    for index, color in color_map.items():
                        if index in schedule_text:
                            label.setStyleSheet(f"background-color: {color.name()};")
                            break

                day += 1

    # 保存 PNG
    def on_save_png(self):
        if not self.current_schedule_dict:
            QMessageBox.warning(self, "提示", "无排班表可保存")
            return

        month = self.month_combo.currentData()
        original_text = self.drag_label.text()
        name = self.name_input.text()
        self.drag_label.setText(f"{name}的{month}月排班表")
        scale_factor = 2

        w = self.width() * scale_factor
        h = self.height() * scale_factor
        pixmap = QPixmap(w, h)
        pixmap.setDevicePixelRatio(scale_factor)
        pixmap.fill(Qt.white)
        self.render(pixmap)

        img = pixmap.toImage()
        rect = img.rect()
        top = bottom = left = right = 0

        # 自动裁掉空白
        for y in range(rect.height()):
            for x in range(rect.width()):
                if QColor(img.pixel(x, y)) != QColor(Qt.white):
                    top = y
                    break
            else:
                continue
            break
        for y in reversed(range(rect.height())):
            for x in range(rect.width()):
                if QColor(img.pixel(x, y)) != QColor(Qt.white):
                    bottom = y
                    break
            else:
                continue
            break
        for x in range(rect.width()):
            for y in range(rect.height()):
                if QColor(img.pixel(x, y)) != QColor(Qt.white):
                    left = x
                    break
            else:
                continue
            break
        for x in reversed(range(rect.width())):
            for y in range(rect.height()):
                if QColor(img.pixel(x, y)) != QColor(Qt.white):
                    right = x
                    break
            else:
                continue
            break

        cropped_pixmap = pixmap.copy(QRect(left, top, right - left + 1, bottom - top + 1))
        name = self.name_input.text()
        save_path = os.path.join(os.getcwd(), f"{name}的{month}月排班表.png")
        cropped_pixmap.save(save_path, "PNG")
        self.drag_label.setText(original_text)
        QMessageBox.information(self, "提示", f"排班表已保存为: {save_path}")

# ------------------- 主程序 -------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ScheduleApp()
    window.show()
    sys.exit(app.exec_())
