import pandas as pd
import os, math

import matplotlib.font_manager as fm
import matplotlib.pyplot as plt
from pylab import mpl

class NessusImageGenerater:
    def __init__(self):
        pass

    def generate_image(self, df, project_path, img_name):
        # font_path = "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"  # 你的字體路徑
        # plt.rcParams["font.sans-serif"] = fm.FontProperties(fname=font_path).get_name()
        mpl.rcParams["font.sans-serif"] = ["DFKai-SB"]  
        # 標楷體： DFKai-SB        # 微軟正黑體： Microsoft JhengHei
        mpl.rcParams["axes.unicode_minus"] = False
        plt.rcParams['axes.unicode_minus'] = False  # 避免負號顯示錯誤
        
        x = df.columns
        y = df.iloc[0]

        plt.figure(figsize=(8, 5))
        plt.bar(x, y, color=['red', 'orange', 'yellow', 'blue'])

        # 設定標籤和標題
        plt.xlabel("弱點等級")
        plt.ylabel("數量")
        # plt.title("Bar Chart from Existing DataFrame")

        y_max = max(y) * 1.5
        plt.ylim(0, math.ceil(y_max))
        ## 確保 y 軸刻度為整數
        if max(y) < 4:
            plt.yticks(range(0, int(max(y) * 1.5) + 1))  # 只顯示整數刻度
        # 顯示數據標籤
        high = max(y)*0.05
        for i, v in enumerate(df.iloc[0]):
            plt.text(i, v+high, f"{v}", ha='center', fontsize=12)
        
        project_path = os.path.join(project_path, img_name)
        plt.savefig(project_path, dpi=300, bbox_inches='tight')
        plt.close()  # 關閉圖表，防止多餘的記憶體佔用

    def image_data(self, df):
        all_severities = ["4", "3", "2", "1"]  # 定義所有可能的級別
        df_count = df['severity'].value_counts().reindex(all_severities, fill_value=0).sort_index(ascending=False).to_frame().T
        host_count = df.groupby("severity")["ip"].nunique().reindex(all_severities, fill_value=0).sort_index(ascending=False).to_frame().T

        severity_mapping = {"4": "嚴重", "3": "高", "2": "中", "1": "低"}
        df_count.columns = [f"{severity_mapping[col]}" if col in severity_mapping else col for col in df_count.columns]
        host_count.columns = [f"{severity_mapping[col]}" if col in severity_mapping else col for col in host_count.columns]
        return df_count, host_count

        
def main():
    # 先計算所有級別的數量，並確保所有級別都出現
    all_severities = ["4", "3", "2", "1"]  # 定義所有可能的級別
    df_count = df['severity'].value_counts().reindex(all_severities, fill_value=0).sort_index(ascending=False).to_frame().T
    
    

    severity_mapping = {"4": "嚴重", "3": "高", "2": "中", "1": "低"}
    df_count.columns = [f"{severity_mapping[col]}" if col in severity_mapping else col for col in df_count.columns]
    
    
    i_g = NessusImageGenerater()
    i_g.generate_image(df_count)
    pass

if __name__ == "__main__":
    main()