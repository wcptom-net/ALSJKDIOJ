import os
import requests
import pandas as pd
from xml.etree import ElementTree as ET
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QLabel, QWidget
from PyQt5.QtCore import Qt

# API 配置
API_URL = "https://dst.apigateway.data.gov.mo/dst_hotel"
HEADERS = {"Authorization": "APPCODE 09d43a591fba407fb862412970667de4"}

# 發送請求並解析數據
response = requests.get(API_URL, headers=HEADERS)
response.raise_for_status()
root = ET.fromstring(response.content)
data = []
for hotel in root.findall(".//hotel"):
    data.append({
        "id": hotel.findtext("id"),
        "classname_zh": hotel.findtext("classname_zh"),
        "latitude": hotel.findtext("latitude"),
        "longitude": hotel.findtext("longitude"),
        "green_hotel": hotel.findtext("green_hotel"),
        "room_no": int(hotel.findtext("room_no") or 0),
        "name_zh": hotel.findtext("name_zh"),
        "address_zh": hotel.findtext("address_zh"),
    })

# 轉換為 DataFrame
df = pd.DataFrame(data)

# 將 green_hotel 列的值替換為 "是" 或 "否"
df["green_hotel"] = df["green_hotel"].apply(lambda x: "是" if x == "1" else "否")

# 計算統計數據
total_hotels = len(df)
total_green_hotels = len(df[df["green_hotel"] == "是"])

# 獲取當前 Python 檔案的目錄
current_dir = os.path.dirname(os.path.abspath(__file__))

# 儲存數據到 Excel
def save_files():
    # 按地址分類
    def classify_address(address):
        if not address:  # 檢查是否為 None 或空值
            return "其他"
        if "澳門" in address:
            return "澳門"
        elif "氹仔" in address:
            return "氹仔"
        elif "路氹" in address:
            return "路氹"
        elif "路環" in address:
            return "路環"
        else:
            return "其他"

    df["address_group"] = df["address_zh"].apply(classify_address)
    address_groups = df.groupby("address_group").size().reset_index(name="count")
    address_groups.to_excel(os.path.join(current_dir, "hotels_address_groups.xlsx"), index=False)

    # 按 classname_zh 分類的酒店數量
    class_groups = df.groupby("classname_zh").size().reset_index(name="count")
    class_groups.to_excel(os.path.join(current_dir, "hotels_class_groups.xlsx"), index=False)

    # 按 classname_zh 分類的酒店房間總數
    class_total_rooms = df.groupby("classname_zh")["room_no"].sum().reset_index(name="total_rooms")
    class_total_rooms.to_excel(os.path.join(current_dir, "hotels_class_total_rooms_groups.xlsx"), index=False)

    # 整個大表按房間數量倒序排序
    sorted_df = df.sort_values(by="room_no", ascending=False)
    sorted_df.to_excel(os.path.join(current_dir, "hotels.xlsx"), index=False, columns=["id", "classname_zh", "latitude", "longitude", "green_hotel", "room_no", "name_zh", "address_zh"])

# 在背景儲存文件
save_files()

# 創建 GUI
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("酒店數據統計")
        self.setGeometry(100, 100, 500, 300)

        # 主窗口佈局
        layout = QVBoxLayout()

        # 顯示總酒店數量
        self.total_hotels_label = QLabel(f"總共有 {total_hotels} 間酒店。")
        self.total_hotels_label.setAlignment(Qt.AlignCenter)
        self.total_hotels_label.setStyleSheet("font-size: 18px; font-weight: bold; color: #333;")
        layout.addWidget(self.total_hotels_label)

        # 顯示獲得環保酒店獎的數量
        self.green_hotels_label = QLabel(f"總共有 {total_green_hotels} 間獲得環保酒店獎。")
        self.green_hotels_label.setAlignment(Qt.AlignCenter)
        self.green_hotels_label.setStyleSheet("font-size: 18px; font-weight: bold; color: #4CAF50;")
        layout.addWidget(self.green_hotels_label)

        # 提示文件已儲存
        self.file_saved_label = QLabel("數據已自動儲存為以下文件：\n"
                                       "1. hotels_address_groups.xlsx\n"
                                       "2. hotels_class_groups.xlsx\n"
                                       "3. hotels_class_total_rooms_groups.xlsx\n"
                                       "4. hotels.xlsx")
        self.file_saved_label.setAlignment(Qt.AlignCenter)
        self.file_saved_label.setStyleSheet("font-size: 14px; color: #555;")
        layout.addWidget(self.file_saved_label)

        # 設置中心窗口
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        # 設置窗口樣式
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f9f9f9;
            }
            QLabel {
                margin: 10px;
            }
        """)


if __name__ == "__main__":
    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec_()