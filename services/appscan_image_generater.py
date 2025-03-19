import pandas as pd
import os, json

# import matplotlib.font_manager as fm
import matplotlib.pyplot as plt
from pylab import mpl
import math

class AppScanImageGenerater:
# 假設 df 已經存在，並且只有一行數據
    def __init__(self):
        print("AppScanImageGenerater --- loaded")
        # # 包含參考資訊
        # self.bar_color = ['gray', 'yellow', 'orange', '#FF5151', '#CE0000']
        # self.severity_mapping = {"Critical": "嚴重", "High": "高", "Medium": "中", "Low": "低", "Informational": "參考資訊"}
        # 不包含參考資訊
        self.bar_color = ['yellow', 'orange', '#FF5151', '#CE0000']
        self.severity_mapping = {"Critical": "嚴重", "High": "高", "Medium": "中", "Low": "低"}
        self.severity_order = ['嚴重', '高', '中', '低', '參考資訊']
    
    def _cause_count(self, data, need_canse):   # 計算嚴重程度數量、排除參考資料、轉換英文標籤為中文
        cause_count = pd.DataFrame({'severity': self.severity_order, 'amount': [0] * len(self.severity_order)})   # 初始化
        cause_count = pd.concat([cause_count, data]).groupby("severity", as_index=False).sum()
        cause_count["severity"] = pd.Categorical(cause_count["severity"], categories=self.severity_order, ordered=True)
        cause_count = cause_count.sort_values("severity").reset_index(drop=True)
        # 排除參考資料、轉換英文標籤為中文
        cause_count = cause_count.iloc[:need_canse]   # 留嚴重、高、中
        cause_count = cause_count.sort_index(ascending=False).T
        cause_count.columns = [f"{self.severity_mapping[col]}" if col in self.severity_mapping else col for col in cause_count.columns]
        return cause_count

    def _image_start(self, cause_count, task_folder_path, file_name):
        x = cause_count.iloc[0]
        y = cause_count.iloc[1]
        
        ## 設定圖片樣式
        plt.figure(figsize=(8, 5))
        plt.bar(x, y, color=self.bar_color)
        ## 設定標籤和標題
        plt.xlabel("弱點等級")
        plt.ylabel("數量")
        # plt.title("Bar Chart from Existing DataFrame")
        y_max = max(y) * 1.5
        plt.ylim(0, math.ceil(y_max))
        ## 確保 y 軸刻度為整數
        if max(y) < 4:
            plt.yticks(range(0, int(max(y) * 1.5) + 1))  # 只顯示整數刻度
        ## 顯示數據標籤
        high = max(y)*0.05
        for i, v in enumerate(y): 
            plt.text(i, v+high, f"{v}", ha='center', fontsize=12)
        plt.savefig(f"{os.path.join(task_folder_path, "image", file_name)}.jpg", dpi=300, bbox_inches='tight')
        plt.close()  ## 關閉圖表，防止多餘的記憶體佔用

    def generate_image(self, task_folder_path, need_canse=3):
        # font_path = "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"  # 你的字體路徑
        # plt.rcParams["font.sans-serif"] = fm.FontProperties(fname=font_path).get_name()
        mpl.rcParams["font.sans-serif"] = ["DFKai-SB"]  
        # 標楷體： DFKai-SB        # 微軟正黑體： Microsoft JhengHei
        mpl.rcParams["axes.unicode_minus"] = False
        plt.rcParams['axes.unicode_minus'] = False  # 避免負號顯示錯誤

        json_paths = []
        for root, _, files in os.walk(os.path.join(task_folder_path, "json")):
            json_paths.extend([os.path.join(root, f) for f in files if f.endswith('.json')])

        total_count = pd.DataFrame({'severity': self.severity_order, 'amount': [0] * len(self.severity_order)})   # 初始化
        for json_path in json_paths:
            file_name = os.path.splitext(os.path.basename(json_path))[0]
            data = self._to_df(json_path)["risk_web"]
            # 計算嚴重程度數量
            count = data.groupby("severity").size().reset_index(name='amount')
            cause_count = self._cause_count(count, need_canse)
            # 輸出圖表
            self._image_start(cause_count, task_folder_path, file_name)
            # 統計總數量
            total_count = pd.concat([total_count, count]).groupby("severity", as_index=False).sum()
        # 輸出總數量的圖表
        total_cause_count = self._cause_count(total_count, need_canse)
        # 輸出圖表
        file_name = os.path.basename(os.path.dirname(task_folder_path))
        self._image_start(total_cause_count, task_folder_path, file_name)

    @staticmethod
    def _to_df(json_path):
        # 讀取 JSON 檔案
        with open(json_path, "r", encoding="utf-8") as f:
            json_data = json.load(f)

        # 轉回原本的 dict[dict[DataFrame]]
        dfs = {key: pd.DataFrame(records) for key, records in json_data.items()}
        return dfs
        
def main():
    # 先計算所有級別的數量，並確保所有級別都出現
    # df = pd.DataFrame({
    #     'severity': ['嚴重', '高', '中', '低', '參考資訊'],
    #     'amount': [2, 7, 14, 28, 3]
    # })
    task_folder_path = "D:/GitHub/AppScan_report/data/uploadproject/test01"
    
    need_canse=3  # 只需要幾個風險填幾個
    ScanImageGenerater = AppScanImageGenerater()
    ScanImageGenerater.generate_image(task_folder_path, need_canse=need_canse)
    print("OK")

if __name__ == "__main__":
    main()