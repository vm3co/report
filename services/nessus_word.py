# std
import os, json
import pandas as pd
from datetime import datetime
import time

# doc
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm  # 可用於設置圖片尺寸

# from services.word_refresh import MenuUpdater

class NessusDocGenerater:
    def __init__(self):
        print("NessusDocGenerater --- loaded")
        # 加載模板文件
        self.template = DocxTemplate(os.path.join(os.getcwd(), "data/report_templates","template_va.docx"))
        # self.menu_updater = MenuUpdater()
        # self.template = DocxTemplate("template.docx")

    def process_table_1(self, df, context):
        # table_1
        table = df[["start", "end", "target", "count"]].copy()
        context["table_tar_info"] = table.to_dict(orient="records")

        return context

    def process_table_2(self, df, context):
        # table_2
        table = df.groupby("severity").size().reset_index(name="count")
        summary_list = [
            {
                "crit": table[table["severity"] == "4"]["count"].sum(),
                "high": table[table["severity"] == "3"]["count"].sum(),
                "med": table[table["severity"] == "2"]["count"].sum(),
                "low": table[table["severity"] == "1"]["count"].sum(),
                "total": df.shape[0],
            }
        ]
        host_table = df.groupby("severity")["ip"].nunique().reset_index(name="count")
        host_summary_list = [
            {
                "host_crit": host_table[host_table["severity"] == "4"]["count"].sum(),
                "host_high": host_table[host_table["severity"] == "3"]["count"].sum(),
                "host_med": host_table[host_table["severity"] == "2"]["count"].sum(),
                "host_low": host_table[host_table["severity"] == "1"]["count"].sum(),
            }
        ]

        context["table_va_summary"] = summary_list
        context["table_va_summary_host"] = host_summary_list

        return context
    
    def process_image(self, project_path, context):
        image_path = os.path.join(project_path, "bar_summary_va.jpg")
        image = InlineImage(self.template, image_path, width=Mm(120))  # 圖片寬度為
        context["summary_image"] = image
        
        image_host_path = os.path.join(project_path, "summary_host_va.jpg")
        image_host = InlineImage(self.template, image_host_path, width=Mm(120))  # 圖片寬度為
        context["summary_host_va"] = image_host

        return context

    def process_table_3(self, df, context):
        # table_3
        table = df.groupby(["file_no", "ip"]).agg(
            crit=("severity", lambda x: (x == "4").sum()),
            high=("severity", lambda x: (x == "3").sum()),
            med=("severity", lambda x: (x == "2").sum()),
            low=("severity", lambda x: (x == "1").sum()),
            total=("severity", "count"),
        ).reset_index().drop(columns=["file_no"])

        # 添加行號
        table.insert(0, "no", range(1, len(table) + 1))

        context["table_va_count"] = table.to_dict(orient="records")

        return context

    def process_table_4(self, df, context):
        # table_4
        table = df.groupby(["severity", "pluginid", "pluginname"]).size().reset_index(name="count")
        table = table.sort_values(by=["severity", "count"], ascending=[False, False])
        
        severity_mapping = {4: "高", 3: "中", 2: "低", 1: "低"}
        table["severity"] = table["severity"].astype(int).map(severity_mapping)
        
        # 添加行號
        table.insert(0, "no", range(1, len(table) + 1))  

        table = table.rename(columns={
            "severity": "risk",
            "pluginid": "pluginnum",
            "pluginname": "pluginname"
         })

        context["table_va_sum"] = table.to_dict(orient="records")

        return context

    def set_common_data(self, title_word, context):
        # common_data
        # project_name1 = "此報告為"
        # project_name2 = "系統自動產生"
        # project_name3 = "請手動填入標題頁首及文編"

        project_name1 = title_word["company_name"]
        project_name2 = title_word["project_name"]
        project_name3 = title_word["report_name"]

        # header1 = f"{p1} {p2}"
        # header2 = f"{p3}"
        header_line1 = f"{project_name1} {project_name2}"
        header_line2 = f"{project_name3}"

        file_no = title_word["file_no"]

        context.update({
            "project_name1": f"{project_name1}",
            "project_name2": f"{project_name2}",
            "project_name3": f"{project_name3}",

            "header_line1": f"{header_line1}",
            "header_line2": f"{header_line2}",

            "file_no": f"{file_no}",
        })

        return context

    def set_date(self, df, context):
        now = datetime.now()

        # 格式化日期
        date_dot = now.strftime("%Y.%m.%d")
        date_ch = now.strftime("%Y年%m月%d日")

        context.update({
            "date_dot": f"{date_dot}",
            "date_ch": f"{date_ch}",
        })
        return context
    
    def set_summary_word(self, df, context):
        # 計算總數
        number_total = len(df)
        
        # 計算 severity = 4 的數量
        number_crit = df[df["severity"] == "4"].shape[0]
        
        # 找出出現最多次的 pluginname
        va_summary = df["pluginname"].mode()[0] if not df["pluginname"].empty else None
        
        context.update({
            "number_total": f"{number_total}",
            "number_crit": f"{number_crit}",
            "va_summary": f"{va_summary}",
        })
        return context


    def generate_report(self, title_word, project_path):
        output_path = os.path.join(project_path, "outputfile.docx")

        information = self._to_df(os.path.join(project_path, "nessus.json"))
        df_nessus, df_scan_info = information["df_nessus"], information["df_nessus_scan_info"]
        
        df_nessus = df_nessus[df_nessus['severity'] != "0"]
        # init context
        context = {}
        
        # 設置通用信息
        context = self.set_common_data(title_word, context)
        context = self.set_date(df_nessus, context)
        context = self.set_summary_word(df_nessus, context)

        # 依次處理各個表格
        context = self.process_table_1(df_scan_info, context)
        context = self.process_table_2(df_nessus, context)
        context = self.process_image(project_path, context)
        context = self.process_table_3(df_nessus, context)
        context = self.process_table_4(df_nessus, context)

        # 渲染模板並輸出報告
        self.template.render(context)
        self.template.save(output_path)
        
        # 更新目錄
        # self.menu_updater.re_index(output_path)

        print(f"Nessus_Word --- fin")

    @staticmethod
    def _to_df(json_path):
        # 讀取 JSON 檔案
        with open(json_path, "r", encoding="utf-8") as f:
            json_data = json.load(f)

        # 轉回原本的 dict[dict[DataFrame]]
        dfs = {key: pd.DataFrame(records) for key, records in json_data.items()}
        return dfs


def main():
    project_path = r"D:\GitHub\appliance\app\data\uploadproject\TEST3"
    temp = NessusDocGenerater()
    temp.generate_report(df_nessus, df_scan_info, project_path)

if __name__ == "__main__":
    main()

