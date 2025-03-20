import requests
from bs4 import BeautifulSoup

import pandas as pd
import os

class ApplinaceTranslate:
    def __init__(self):
        self.columns = ["pluginId", "pluginName", "description", "solution"]
        self.trans_data = None
        self.file_name = "trans_db.xlsx"
        self.init_db_xlsx()
        
    def init_db_xlsx(self):
        current_path = os.getcwd()  # 取得當前檔案目錄
        data_folder = os.path.join(current_path, "data/trans_data")  # 定義相對於檔案的目標路徑
        os.makedirs(data_folder, exist_ok=True)
        
        self.file_name = os.path.join(data_folder, "trans_db.xlsx")
        
        if not os.path.exists(self.file_name):
            df = pd.DataFrame(columns=self.columns)
            df.to_excel(self.file_name, index=False)

        self.trans_data = pd.read_excel(self.file_name, sheet_name=0)

    def nessus_zh_tw(self, plugin_id:int):
        data = {}
        url = f"https://zh-tw.tenable.com/plugins/nessus/{plugin_id}" 
        response = requests.get(url)
        
        if response.status_code != 200:
            return None
        
        response.encoding = 'utf-8'  
        soup = BeautifulSoup(response.text, 'html.parser')
        
        plugin_title = soup.find('h1', class_='h2')
        plugin_title = plugin_title.get_text(strip=True)
        data.update({'pluginName' : f'{plugin_title}'})

        sections = soup.find_all('section', class_='mb-3')
        if sections:
            for section_idx, section in enumerate(sections, start=1):
                sub_title = section.find('h4', class_ ='border-bottom pb-1')
                sub_title = sub_title.get_text(strip=True)
                
                span_texts = section.find_all('span')
                content_full = []
                for span_idx, span in enumerate(span_texts, start=1):
                    content = span.get_text(strip=True)
                    if content:
                        content_full.append(content)

                if not span_texts:
                    p_texts = section.find_all('p')
                    for p_idx, p in enumerate(p_texts, start=1):
                        content = p.get_text(strip=True)
                        content_full.append(content)
                
                content_full = "\n".join(content_full)
                data.update({f'{sub_title}' : content_full})
        else:
            pass
        
        values = list(data.values())
        selected_values = [plugin_id] + [values[i] for i in [0, 2, 3]]
        data = dict(zip(self.columns, selected_values))
        return data

    def xlsx_update(self, data):
        if not data:
            return
            
        new_df = pd.DataFrame([data])
        
        self.trans_data = pd.concat([self.trans_data, new_df], ignore_index=True)
        self.trans_data.to_excel(self.file_name, index=False)

    def db_search(self, plugin_id):
        df = self.trans_data
        if df.index.name != "pluginId":
            df = df.set_index("pluginId")
        
        if plugin_id in df.index:
            return df.loc[plugin_id].to_dict()
        else:
            return None
    
    def trans_run(self,plugin_id):
        plugin_id = int(plugin_id)
        
        db_data = self.db_search(plugin_id)
        if not db_data:
            db_data = self.nessus_zh_tw(plugin_id)
            self.xlsx_update(db_data)

        return db_data
    
def main():
    plugin_id = 44330
    trans = ApplinaceTranslate()
    data = trans.trans_run(plugin_id)

if __name__ == "__main__":
    main()
