import os, json
import pandas as pd
from datetime import datetime
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm
# from services.appscan_parser import AppScanParser

class AppScanWordReport:
    def __init__(self):
        self.template_path = os.path.join(os.getcwd(), "data/report_templates", "template_web.docx")
        self.version_image_path = os.path.join(os.getcwd(), "data/img")
        # self.template_path = os.path.join("D:/GitHub/report", "data/report_templates", "template_web.docx")
        # self.version_image_path = os.path.join("D:/GitHub/report", "data/img")
        # self.severity_order = ['嚴重', '高', '中', '低', '參考資訊']
        self.severity_order = ['嚴重', '高', '中']


    def _version_img(self, context, df_logs, doc):
        version_name = df_logs[0]["version_web"]
        appscan_version_image = InlineImage(doc, os.path.join(self.version_image_path, f"{version_name}.png"), width=Cm(12.6))
        context["appscan_version_image"] = appscan_version_image
        
    def _df_log(self, index, df):
        log_dict = {
            "index":index+1, 
            "web_url": df["information"].loc[0, "web_url"], 
            "start_time": df["information"].loc[0, "start_time"],
            "version_web": df["information"].loc[0, "version_web"]
            }
        return log_dict

    def _web_causes(self, df_risk_web):
        web_causes = df_risk_web.groupby(['name', 'risk', 'cause', 'severity',
                                          'solution_description', 'solution_suggest']).size().reset_index(name='amount')
        
        web_causes["severity"] = pd.Categorical(web_causes["severity"], categories=self.severity_order, ordered=True)
        web_causes = web_causes.sort_values("severity").reset_index(drop=True)
        return web_causes

    @staticmethod
    def _total_counts(counts, web_causes, severity_order):
        cause_count = web_causes.groupby("severity", observed=True)['amount'].sum().reset_index()

        for sev in severity_order:
            # 使用 .loc[] 來篩選 'severity' 欄位等於 sev 的行，然後取 'amount' 的值
            amount_value = cause_count.loc[cause_count["severity"] == sev, "amount"]
            # 如果篩選後 DataFrame 不是空的，就取出第一個值，否則預設為 0
            counts[sev] += amount_value.iloc[0] if not amount_value.empty else 0
        return counts
        
    @staticmethod
    def _to_df(json_path):
        # 讀取 JSON 檔案
        with open(json_path, "r", encoding="utf-8") as f:
            json_data = json.load(f)

        # 轉回原本的 dict[dict[DataFrame]]
        dfs = {key: pd.DataFrame(records) for key, records in json_data.items()}
        return dfs
    
    def _generate_vulnerability_summary(self, df_risk_web, web_summary, web_causes_list, index):        
        web_causes = self._web_causes(df_risk_web)
        # 篩選出 "嚴重" 的弱點
        severe_issues = web_causes.loc[web_causes["severity"] == "嚴重", "cause"]
        if web_summary == "" and not severe_issues.empty:
            web_summary = f"主要為{severe_issues.iloc[0]}等弱點，"
        # 計算風險數量(用在表格)
        counts = {sev: 0 for sev in self.severity_order[:3]}
        counts = self._total_counts(counts, web_causes, self.severity_order[:3])
        # 弱點說明與修正建議
        web_causes_explanation = df_risk_web.groupby(['severity', 'name', 'risk', 'solution_description', 'solution_suggest']
                                                     ).size().reset_index(name='amount')
        web_causes_explanation["severity"] = pd.Categorical(web_causes_explanation["severity"], categories=self.severity_order, ordered=True)
        web_causes_explanation = web_causes_explanation.sort_values("severity").reset_index(drop=True)
        web_causes_explanation_dict = web_causes_explanation.to_dict(orient='records')        
        # 弱點檢測分析報告
        web_causes_need_dict = web_causes[web_causes["severity"].isin(self.severity_order[:3])].to_dict(orient='records')
        web_causes_list.append({
            "num_web": index+1,
            "web_causes_explanation_dict": web_causes_explanation_dict,
            "web_causes_need_dict": web_causes_need_dict,
            "information_web_url": df_risk_web.loc[0,"goal_url"],
            "count_C_web": counts['嚴重'],
            "count_H_web": counts['高'],
            "count_M_web": counts['中'],
        })
        
        return web_causes_list, web_summary, counts
    
    def _word(self, task_folder_path, json_paths, title_word, need_canse, img_name):
        output_path = os.path.join(task_folder_path, f"{img_name}.docx")
        doc = DocxTemplate(self.template_path)
        context = {key: title_word.get(key, "") for key in ["company_name", "project_name", "report_name", "file_no", "scanner_ip", "date_start", "date_end"]}
        
        df_logs = []
        web_summary = ""
        web_causes_list = []
        total_counts = {sev: 0 for sev in self.severity_order}
        for index, json_path in enumerate(json_paths):
            df = self._to_df(json_path)
            df_logs.append(self._df_log(index, df))
            
            df_risk_web = df["risk_web"]   
            # 留需要的風險程度
            df_risk_web = df_risk_web[df_risk_web["severity"].isin(self.severity_order[:need_canse])]
            web_causes_list, web_summary, counts = self._generate_vulnerability_summary(df_risk_web, web_summary, web_causes_list, index)
            for severity in self.severity_order:
                total_counts[severity] = total_counts[severity] + counts[severity]
            if index == 0:
                self._version_img(context, df_logs, doc)  # appscan的版本
    
        context = self._set_date(context)  # 報告日期
        img_path = os.path.join(task_folder_path, "image", f"{img_name}.jpg")
        context.update({
            "web_causes": web_causes_list,
            "web_summary": web_summary,
            "count_C_web_total": total_counts['嚴重'],
            "count_H_web_total": total_counts['高'],
            "count_M_web_total": total_counts['中'],
            "number_total_web": sum(total_counts.values()),
            "img_web_count_image": InlineImage(doc, img_path, width=Cm(12)),
            "web_goal_lists": df_logs
        })
        
        doc.render(context)
        
        doc.save(output_path)
        print(f"Word 報告已成功生成：{output_path}")

    def generate_word_report(self, title_word, task_folder_path, report_type, need_canse=3):
        json_paths = []
        file_names = []
        for root, _, files in os.walk(task_folder_path):
            json_paths.extend([os.path.join(root, f) for f in files if f.endswith('.json')])
            file_names.extend([os.path.splitext(f)[0] for f in files if f.endswith('.json')])
   
        if report_type == "single":
            img_name = os.path.basename(os.path.dirname(task_folder_path))
            self._word(task_folder_path, json_paths, title_word, need_canse, img_name)
        elif report_type == "multiple":
            for jj, ff in zip(json_paths, file_names):
                img_name = f"{ff}.jpg"
                jj = [jj]
                self._word(task_folder_path, jj, title_word, need_canse, img_name)

    @staticmethod
    def _set_date(context):
        now = datetime.now()
        date_dot, date_ch = now.strftime("%Y.%m.%d"), now.strftime("%Y年%m月%d日")
        context.update({"date_dot": date_dot, "date_ch": date_ch})
        return context


#===================================================================
def test():
    task_folder_path = r"D:\GitHub\report\data\uploadproject\MyTaskName6\appscan"
    # appscan_files = ("Appscan_https___ginandjuice.shop_.scan", "Appscan_http___testasp.vulnweb.com_.scan")
    
    title_word = {
        "company_name": "公司名稱",
        "project_name": "專案名稱",
        "report_name": "報告名稱",
        "file_no": "文編預留位置",
        "scanner_ip": "scanner_ip",
        "date_start": "date_start",
        "date_end": "date_end"
    }
    report_type = "single"  # 報告為單一(single)或多個(multiple)
    # parser = AppScanParser()
    # parser.run("C:/Program Files (x86)/HCL/AppScan Standard/AppScanCMD.exe", appscan_files, task_folder_path, test=True)
    need_canse=3  # 只需要幾個風險填幾個
    report_generator_word = AppScanWordReport()
    report_generator_word.generate_word_report(title_word, task_folder_path, report_type, need_canse)
    
if __name__ == "__main__":
    test()
