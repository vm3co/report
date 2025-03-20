import subprocess, math
from bs4 import BeautifulSoup
import pandas as pd
import os, re, json, shutil
from concurrent.futures import ThreadPoolExecutor

class AppScanParser:
    def __init__(self):
        pass
    
    def _generate_html_batch(self, task_folder_path, appscanCMDexe, batch_id, batch_appscan):   # 產生單個 BAT 檔案並執行
        bat_path = os.path.join(task_folder_path, "bat", f'batch_{batch_id}.bat')
        try:
            with open(bat_path, 'w') as fw:
                fw.write('chcp 950\n')  # 設定編碼
                for appscan in batch_appscan:
                    html_path = os.path.join(task_folder_path, "html", os.path.splitext(os.path.basename(appscan))[0]) + ".html"
                    fw.write(f'"{appscanCMDexe}" report /b "{appscan}" /rf "{html_path}" /rt html\n')            
            # 執行批次檔
            subprocess.run([bat_path], shell=True, check=True)
        except subprocess.CalledProcessError as e:
            print(f"執行 {bat_path} 時發生錯誤: {e}")
                
    def _generate_html(self, appscanCMDexe, task_folder_path, appscan_paths, max_batches=4):   ## 將bat檔分批處理並平行運行
        # 初始化 batches，确保有 max_batches 个空批次
        batches = [(i, []) for i in range(max_batches)]
        
        # 均匀分配任务到不同的 bat 文件
        for index, appscan_path in enumerate(appscan_paths):
            batch_id = index % max_batches  # 轮流分配
            batches[batch_id][1].append(appscan_path)
        
        # 使用 ThreadPoolExecutor 平行執行批次
        with ThreadPoolExecutor(max_workers=max_batches) as executor:
            future_tasks = []
            for batch_id, batch_appscan in batches:
                future = executor.submit(self._generate_html_batch, task_folder_path, appscanCMDexe, batch_id, batch_appscan)
                future_tasks.append(future)            
            # **這裡確保所有批次完成**
            for future in future_tasks:
                future.result()  # 這行會等待對應的批次完成
        ## 刪除bat檔
        shutil.rmtree(os.path.join(task_folder_path, "bat"))

    def _parse_json(self, information_web, json_path):
        json_data = {key: df.to_dict(orient="records") for key, df in information_web.items()}                        
        # 存成 JSON 檔案
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(json_data, f, indent=4, ensure_ascii=False)

    def _parse_html(self, task_folder_path):    # 解析 HTML 報告 
        html_paths = []
        for root, _, files in os.walk(os.path.join(task_folder_path, "html")):
            html_paths.extend([os.path.join(root, f) for f in files if f.endswith('.html')])

        for html_path in html_paths:
            with open(html_path, encoding="utf-8") as f:
                soup = BeautifulSoup(f, "lxml")            
            version_web = soup.find("body").text.split("掃描開始時間")[0].split("所建立")[1].strip()
            web_url = "https://" + soup.find("div", string="主機").find_next_sibling("div").text.strip()
            start_time = soup.find("div", string="掃描開始時間：").find_next_sibling("div").text.strip()

            information = {
                "version_web": version_web,
                "web_url": web_url,
                "start_time": start_time,
                }
                        
            risk_entries = []
            name_dict = {
                "嚴重性：": "severity",
                "CVSS 評分：": "cvss",
                "CVSS 向量：": "cvss_vector",
                "CVE：": "cve",
                "URL：": "url",
                "實體：": "entity",
                "風險：": "risk",
                "原因：": "cause",
                "修正：": "solution"
            }
            for issue in soup.find_all("div", "issueHeader"):
                entry_keys = ["name", "goal_url", "severity", "cvss", "cvss_vector", "cve", "url", "entity", 
                              "risk", "cause", "solution", "mark", "solution_description", "solution_suggest"]                
                entry = dict.fromkeys(entry_keys, "")
                entry.update({
                    "name"    : issue.find("div", "headerIssueType").text.strip(), 
                    "goal_url": web_url
                    })                
                for name_tag in issue.find_all("div", "name"):
                    value_tag = name_tag.find_next_sibling("div")
                    if name_tag.text == "風險：":
                        des = value_tag.find_all("li")
                        content = ""
                        for cc in des:
                            content+=cc.text
                            content+="\n"
                    elif name_tag.text == "修正：":
                        content = value_tag.text.strip()
                        # 抓修補建議 
                        solution_id = value_tag.a["href"].strip("#")
                        solution_title = soup.find("a", attrs={"name": solution_id}).find_previous()                        
                        solution_description = solution_title.find_next("h4", string="風險：").find_next("ul").text.strip()
                        solution_suggest = solution_title.find_next("h4", string="修正建議：").find_next("ul").text.strip()
                        entry.update({
                            "solution_description": self.character_replace(solution_description).strip(),
                            "solution_suggest"    : self.character_replace(solution_suggest).strip() + '<w:br w:type="page"/>'
                            })
                    else:
                        content = "嚴重" if value_tag.text.strip() == "重大" else value_tag.text.strip()
                    entry[name_dict.get(name_tag.text, name_tag.text)] = content
                
                # 抓紅字
                mark_list = issue.find_next_sibling("div").find_all("span", "mark")
                mark = ""
                for mm in mark_list:
                    mark+=mm.text
                    mark+="\n"
                entry["mark"] = mark.strip()

                risk_entries.append(entry)
                
            risk_df = pd.DataFrame(risk_entries)
 
            file_name = os.path.splitext(os.path.basename(html_path))[0]
            information_web = {
                "risk_web": risk_df,
                "information": pd.DataFrame.from_dict([information])
            }

            json_path = os.path.join(task_folder_path, "json", f"{file_name}.json")
            self._parse_json(information_web, json_path)        


    def run(self, task_folder_path, appscanCMDexe):
        appscan_paths = []
        for root, _, files in os.walk(task_folder_path):
            appscan_paths.extend([os.path.join(root, f) for f in files if f.endswith('.scan')])

        # 製作html、bat、json、image資料夾
        os.makedirs(os.path.join(task_folder_path, "html"), exist_ok=True)
        os.makedirs(os.path.join(task_folder_path, "bat"), exist_ok=True)
        os.makedirs(os.path.join(task_folder_path, "json"), exist_ok=True)
        os.makedirs(os.path.join(task_folder_path, "image"), exist_ok=True)

        # 生成html檔(製作appscanCMDexe的bat檔)
        self._generate_html(appscanCMDexe, task_folder_path, appscan_paths, max_batches=4)

        # 生成json檔
        self._parse_html(task_folder_path)

    @staticmethod
    def character_replace(word):
        """ 替換 word 渲染出錯的字元 """
        if not isinstance(word, str):
            return word
        replacements = {
            "&": "&amp;",
            "<": "&lt;",
            ">": "&gt;"
        }
        word = re.sub(r'\s+', ' ', word).strip()
        for key, value in replacements.items():
            word = word.replace(key, value)
        return word    


def test():
    # appscan_files = [
    #     "Appscan_https___ginandjuice.shop_.scan",
    #     "Appscan_http___testasp.vulnweb.com_.scan"
    # ]
    task_folder_path = "D:/GitHub/AppScan_report/data/uploadproject/test01"
    appscanCMDexe = "C:/Program Files (x86)/HCL/AppScan Standard/AppScanCMD.exe"
    parser = AppScanParser()
    parser.run(task_folder_path, appscanCMDexe)

if __name__ == "__main__":
    test()
