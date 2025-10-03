import sys
import os
import re
import pandas as pd
from PyQt5.QtGui import QColor, QPixmap, QPainter, QFont, QTextDocument
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QPushButton, QTableWidget,
    QTableWidgetItem, QHeaderView, QHBoxLayout, QComboBox, QTextEdit,
    QMessageBox
)
from PyQt5.QtCore import Qt, QDate, QRect

# ------------------- Excel -> TXT -------------------
replace_rules = [
    ('备1', '备1-可休'), ('十四', '8.00~16.45 十四'), ('夜', '21.30~7.30 夜'), ('晚', '15.30~22.30 晚'),
    ('早', '7.30~15.00 早'), ('中', '15.00~21.30 中'), ('四', '8.00~16.45 四'), ('五', '8.00~12.00 五'),
    ('二', '8.00~16.45 二'), ('三', '8.00~16.45 三'), ('九', '8.00~16.45 九'), ('十', '8.00~16.45 十'),
    ('1', '8.00~16.45 1'), ('2', '7.00-15.00 2'), ('3', '7.00-15.00 3'), ('4', '7.30~17.15 4'),
    ('5', '8.00~17.15 5'), ('6', '7.30~16.45 6'), ('7', '8.30~17.15 7'), ('9', '7.30~16.55 9'),
    ('15', '8.00~12.00 15'), ('西', '8.00~12.00 西'), ('备', '7.30~16.30 备'), ('帮', '8.30~16.15 帮'),
    ('休', '休息'), ('工', '工休')
]

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
                for target_word, replace_word in replace_rules:
                    columns[column_number - 1] = re.sub(
                        rf'(?<!\S){target_word}(?!\S)', replace_word, columns[column_number - 1]
                    )
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
        columns_to_keep = [col for col in ['日期','星期','章学亭'] if col in df.columns]
        df = df[columns_to_keep]

        temp_txt = os.path.join(output_dir, os.path.basename(file_path).split('.')[0] + '_output.txt')
        df.to_csv(temp_txt, index=False, encoding='utf-8-sig', sep='\t')
        txt_file_paths.append(temp_txt)

    all_modified_lines = []
    for txt_file in txt_file_paths:
        all_modified_lines.extend(replace_column_in_text(txt_file, replace_rules, column_number=3))

    final_output_path = os.path.join(output_dir, f'章总的{month}月排班表.txt')
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
        self.setWindowTitle("附二院排班表生成器 (V1.3.1)")
        self.resize(900, 650)
        layout = QVBoxLayout()

        self.drag_label = QLabel("把 Excel 文件拖到这里，生成排班表")
        self.drag_label.setAlignment(Qt.AlignCenter)
        self.drag_label.setStyleSheet("border: 2px dashed #aaa; font-size: 18px; padding: 10px;")
        layout.addWidget(self.drag_label)

        # 年月选择
        ym_layout = QHBoxLayout()
        self.year_combo = QComboBox()
        current_year = QDate.currentDate().year()
        for y in range(current_year-5, current_year+6):
            self.year_combo.addItem(str(y), y)
        self.year_combo.setCurrentText(str(current_year))

        self.month_combo = QComboBox()
        for m in range(1, 13):
            self.month_combo.addItem(f"{m}月", m)
        self.month_combo.setCurrentIndex(QDate.currentDate().month()-1)

        ym_layout.addWidget(QLabel("年份:"))
        ym_layout.addWidget(self.year_combo)
        ym_layout.addWidget(QLabel("月份:"))
        ym_layout.addWidget(self.month_combo)
        layout.addLayout(ym_layout)

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
        color_map = {"休": QColor("#90EE90"), "夜": QColor("#D3D3D3"), "21.30~7.30": QColor("#D3D3D3")}

        for row in range(total_rows):
            for col in range(7):
                if row == 0 and col < start_col:
                    self.table.setItem(row, col, QTableWidgetItem(""))
                    continue
                if day > days_in_month:
                    self.table.setItem(row, col, QTableWidgetItem(""))
                    continue

                date_str = QDate(year, month, day).toString("yyyy-MM-dd")
                text = f"{day}\n{schedule_dict.get(date_str, '')}"

                item = QTableWidgetItem(text)
                item.setTextAlignment(Qt.AlignTop | Qt.AlignLeft)
                item.setToolTip(text)

                for k, c in color_map.items():
                    if k in text:
                        item.setBackground(c)
                        break

                self.table.setItem(row, col, item)
                day += 1

    # 保存 PNG
    def on_save_png(self):
        if not self.current_schedule_dict:
            QMessageBox.warning(self, "提示", "无排班表可保存")
            return

        month = self.month_combo.currentData()
        original_text = self.drag_label.text()
        self.drag_label.setText(f"章总的{month}月排班表")
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
        save_path = os.path.join(os.getcwd(), f"章总的{month}月排班表.png")
        cropped_pixmap.save(save_path, "PNG")
        self.drag_label.setText(original_text)
        QMessageBox.information(self, "提示", f"排班表已保存为: {save_path}")

# ------------------- 主程序 -------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ScheduleApp()
    window.show()
    sys.exit(app.exec_())

# ------------------- 版本号 -------------------
APP_VERSION = "V1.3.1"
