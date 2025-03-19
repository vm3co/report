import requests

# 伺服器的 API URL
BASE_URL = "http://localhost:8080/download"

# 要傳遞的參數
params = {
    "name": "MyTaskName",
    "company_name": "Tech Corp",
    "project_name": "AI Research",
    "report_name": "Quarterly Analysis",
    "file_no": "A12345",
    "scanner_ip": "192.168.1.100",
    "date_start": "2025-01-01",
    "date_end": "2025-03-12",
    "type": "multiple"  # 可選 "single" 或 "multiple"
}

# 發送 GET 請求
response = requests.get(BASE_URL, params=params)

# 檢查是否成功
if response.status_code == 200:
    # # 嘗試將回應當作檔案下載
    # with open("downloaded_report.zip", "wb") as file:
    #     file.write(response.content)
    print(f"✅ 發送請求成功: {response.text}")
else:
    print(f"❌ 下載失敗，狀態碼: {response.status_code}, 訊息: {response.text}")