import os
import requests
import pandas as pd
from xml.etree import ElementTree as ET

# API 配置
API_URL = "https://dst.apigateway.data.gov.mo/dst_hotel"
HEADERS = {"Authorization": "APPCODE 09d43a591fba407fb862412970667de4"}

# 發送請求
response = requests.get(API_URL, headers=HEADERS)
response.raise_for_status()  # 確保請求成功

# 解析 XML
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

# 總共有多少間酒店
total_hotels = len(df)
print(f"總共有 {total_hotels} 間酒店。")

# 總共有多少間獲得環保酒店獎
green_hotels = df[df["green_hotel"] == "是"]
total_green_hotels = len(green_hotels)
print(f"總共有 {total_green_hotels} 間獲得環保酒店獎。")

# 拆解地址分類
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

# 獲取當前 Python 檔案的目錄
current_dir = os.path.dirname(os.path.abspath(__file__))

# 按地址分類
df["address_group"] = df["address_zh"].apply(classify_address)
address_groups = df.groupby("address_group").size().reset_index(name="count")
address_groups.to_excel(os.path.join(current_dir, "hotels_address_groups.xlsx"), index=False)
print("按地址分類的酒店數量已儲存為 hotels_address_groups.xlsx。")

# 按 classname_zh 分類的酒店數量
class_groups = df.groupby("classname_zh").size().reset_index(name="count")
class_groups.to_excel(os.path.join(current_dir, "hotels_class_groups.xlsx"), index=False)
print("按 classname_zh 分類的酒店數量已儲存為 hotels_class_groups.xlsx。")

# 按 classname_zh 分類的酒店房間總數
class_total_rooms = df.groupby("classname_zh")["room_no"].sum().reset_index(name="total_rooms")
class_total_rooms.to_excel(os.path.join(current_dir, "hotels_class_total_rooms_groups.xlsx"), index=False)
print("按 classname_zh 分類的酒店房間總數已儲存為 hotels_class_total_rooms_groups.xlsx。")

# 整個大表按房間數量倒序排序
sorted_df = df.sort_values(by="room_no", ascending=False)
sorted_df.to_excel(os.path.join(current_dir, "hotels.xlsx"), index=False, columns=["id", "classname_zh", "latitude", "longitude", "green_hotel", "room_no", "name_zh", "address_zh"])
print("整個大表已儲存為 hotels.xlsx。")