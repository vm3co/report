import requests, os

url = "http://localhost:8080/upload/"  # FastAPI 上傳端點


# 建立 `files` 參數 (模擬 JavaScript 的 FormData)
files = []

# 自訂任務名稱
uploadtaskname = "MyTaskName7"
iscontinue = "no"  # 或 "yes"
# 你要上傳的檔案
appscan_files = [
    r"D:\GitHub\AppScan_report\data\uploadproject\test01\Appscan_http___testasp.vulnweb.com_.scan",
    r"D:\GitHub\AppScan_report\data\uploadproject\test01\Appscan_https___ginandjuice.shop_.scan",
    r"D:\GitHub\appliance\app\data\uploadproject\TEST\000_ACSI_1219-1_1(N)_y6zr3g.nessus",
    r"D:\GitHub\appliance\app\data\uploadproject\TEST\000_ACSI_1219-2_1(N)_dtxklw.nessus",
    ]

for appscan_file in appscan_files:
    if not os.path.exists(appscan_file):
        print(f"檔案不存在: {appscan_file}")
# 加入 .appscan 檔案
for appscan_file in appscan_files:
    files.append(("files", (appscan_file, open(appscan_file, "rb"), "application/octet-stream")))

# 加入單一的表單數據
data = {
    "uploadtaskname": uploadtaskname,
    "iscontinue": iscontinue
}

response = requests.post(url, files=files, data=data)


# 顯示回應結果
print(response.json())
