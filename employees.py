import sys
import requests
from openpyxl import Workbook
import os
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QGridLayout, QMessageBox
)
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt

# API 配置
API_URL = "https://dsal.apigateway.data.gov.mo/api/getA2"
API_HEADERS = {
    "Authorization": "APPCODE 09d43a591fba407fb862412970667de4"
}

def fetch_data(year, month):
    """從 API 獲取數據"""
    params = {
        "year": year,
        "month": month
    }
    response = requests.get(API_URL, headers=API_HEADERS, params=params)
    if response.status_code == 200:
        return response.json()
    else:
        raise Exception(f"API 請求失敗，狀態碼: {response.status_code}, 錯誤信息: {response.text}")

def save_to_excel(data, filename):
    """將數據保存到 Excel"""
    wb = Workbook()

    # 第 1 個工作表：按非專業外地僱員人數倒序排序
    ws1 = wb.active
    ws1.title = "非專業外地僱員人數"
    ws1.append(["行業名稱", "細分行業名稱", "行業編號", "企業/實體數目", "非專業外地僱員人數", "家傭外地僱員人數"])
    sorted_data_ne = sorted(data, key=lambda x: x["ne_workers_number"], reverse=True)
    for item in sorted_data_ne:
        ws1.append([
            item["industry_name_tc"],
            item["sub_industry_name_tc"],
            item["industry_code"],
            item["entity_number"],
            item["ne_workers_number"],
            item["xe_workers_number"]
        ])

    # 第 2 個工作表：按專業外地僱員人數倒序排序
    ws2 = wb.create_sheet(title="專業外地僱員人數")
    ws2.append(["行業名稱", "細分行業名稱", "行業編號", "企業/實體數目", "專業外地僱員人數", "家傭外地僱員人數"])
    sorted_data_te = sorted(data, key=lambda x: x["te_workers_number"], reverse=True)
    for item in sorted_data_te:
        ws2.append([
            item["industry_name_tc"],
            item["sub_industry_name_tc"],
            item["industry_code"],
            item["entity_number"],
            item["te_workers_number"],
            item["xe_workers_number"]
        ])

    # 獲取當前腳本所在目錄
    script_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(script_dir, filename)

    # 保存 Excel 檔案
    wb.save(file_path)
    return file_path

def get_latest_available_month():
    """從 API 嘗試獲取最新可用的年份和月份"""
    from datetime import datetime

    now = datetime.now()
    year = now.year
    month = now.month

    while True:
        try:
            params = {"year": year, "month": f"{month:02d}"}
            response = requests.get(API_URL, headers=API_HEADERS, params=params)
            if response.status_code == 200:
                return year, month
            else:
                month -= 1
                if month == 0:
                    month = 12
                    year -= 1
        except Exception:
            return None, None

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        # 嘗試獲取最新可用的年份和月份
        self.latest_year, self.latest_month = get_latest_available_month()

        # 設置窗口標題
        self.setWindowTitle("外僱人數行業統計")
        self.setFixedSize(400, 300)

        # 設置全局字體
        self.setFont(QFont("Arial", 10))

        # 標籤和輸入框
        self.label_info = QLabel(f"最新可用數據：{self.latest_year} 年 {self.latest_month} 月" if self.latest_year and self.latest_month else "無法獲取最新數據")
        self.label_info.setAlignment(Qt.AlignCenter)
        self.label_info.setStyleSheet("font-size: 14px; font-weight: bold; color: #2E86C1;")

        self.label_year = QLabel("年份:")
        self.input_year = QLineEdit(self)
        self.input_year.setText(str(self.latest_year) if self.latest_year else "")
        self.input_year.setPlaceholderText("請輸入年份")
        self.input_year.setStyleSheet("padding: 5px; border: 1px solid #ccc; border-radius: 5px;")

        self.label_month = QLabel("月份:")
        self.input_month = QLineEdit(self)
        self.input_month.setText(f"{self.latest_month:02d}" if self.latest_month else "")
        self.input_month.setPlaceholderText("請輸入月份")
        self.input_month.setStyleSheet("padding: 5px; border: 1px solid #ccc; border-radius: 5px;")

        # 按鈕
        self.button_generate = QPushButton("生成 Excel", self)
        self.button_generate.setStyleSheet("""
            QPushButton {
                background-color: #2E86C1;
                color: white;
                border: none;
                padding: 10px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #1B4F72;
            }
        """)
        self.button_generate.clicked.connect(self.generate_excel)

        # 布局
        layout = QGridLayout()
        layout.addWidget(self.label_info, 0, 0, 1, 2)
        layout.addWidget(self.label_year, 1, 0)
        layout.addWidget(self.input_year, 1, 1)
        layout.addWidget(self.label_month, 2, 0)
        layout.addWidget(self.input_month, 2, 1)
        layout.addWidget(self.button_generate, 3, 0, 1, 2)

        self.setLayout(layout)

    def generate_excel(self):
        """生成 Excel 檔案"""
        year = self.input_year.text()
        month = self.input_month.text()

        # 檢查年份和月份格式
        if not year.isdigit() or not (1900 <= int(year) <= 2100):
            QMessageBox.critical(self, "錯誤", "年份格式不正確！請輸入YYYY格式的年份。")
            return

        if not month.isdigit() or not (1 <= int(month) <= 12):
            QMessageBox.critical(self, "錯誤", "月份格式不正確！請輸入MM格式的月份。")
            return

        try:
            # 調用 API 獲取數據
            data = fetch_data(year, month.zfill(2))  # 確保月份為兩位數格式
            file_path = save_to_excel(data, "employees.xlsx")
            QMessageBox.information(self, "成功", f"Excel 檔案已保存到: {file_path}")
            
            # 成功後關閉介面並結束程式
            self.close()
            QApplication.quit()
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"發生錯誤: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())