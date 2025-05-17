import requests
from openpyxl import Workbook
import os

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
    print(f"Excel 檔案已保存到: {file_path}")

def get_latest_available_month():
    """從 API 嘗試獲取最新可用的年份和月份"""
    from datetime import datetime

    # 獲取當前年份和月份
    now = datetime.now()
    year = now.year
    month = now.month

    while True:
        try:
            # 嘗試調用 API
            params = {"year": year, "month": f"{month:02d}"}
            response = requests.get(API_URL, headers=API_HEADERS, params=params)
            if response.status_code == 200:
                # 如果成功，返回年份和月份
                return year, month
            else:
                # 如果失敗，減少月份
                month -= 1
                if month == 0:
                    month = 12
                    year -= 1
        except Exception as e:
            print(f"發生錯誤: {e}")
            return None, None

def main():
    # 嘗試獲取最新可用的年份和月份
    latest_year, latest_month = get_latest_available_month()
    if (latest_year and latest_month):
        print(f"最新可用數據的年份和月份為: {latest_year} 年 {latest_month} 月")
    else:
        print("無法獲取最新可用數據，請手動輸入年份和月份。")
    
    # 提供用戶選擇年份和月份
    year = input(f"請輸入年份 (例如 {latest_year}): ") if latest_year else input("請輸入年份 (例如 2025): ")
    month = input(f"請輸入月份 (例如 {latest_month}): ") if latest_month else input("請輸入月份 (例如 05): ")
    
    try:
        print("正在從 API 獲取數據...")
        data = fetch_data(year, month)
        print("數據獲取成功，正在生成 Excel 檔案...")
        save_to_excel(data, "employees.xlsx")
        print("Excel 檔案已成功生成: employees.xlsx")
    except Exception as e:
        print(f"發生錯誤: {e}")

if __name__ == "__main__":
    main()