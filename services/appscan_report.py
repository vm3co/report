from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
from datetime import datetime
import os, shutil

from .appscan_parser import AppScanParser
from .appscan_excel import AppScanExcelReport
from .appscan_word import AppScanWordReport
from .appscan_image_generater import AppScanImageGenerater
from .databases import tasklist, comparelist, adminuser, engine


class Web_Report:
    def __init__(self):
        self.Session = sessionmaker(bind=engine)

        self.report_templates_path = os.path.join(os.getcwd(), "report_templates")
        self.ScanParser = AppScanParser()
        self.ScanExcelGenerater = AppScanExcelReport()
        self.ScanWordGenerater = AppScanWordReport()
        self.ScanImageGenerater = AppScanImageGenerater()

    # def _report_generate(self, appscanCMDexe, file_path, word_title, task_folder_path, test):
    #     json_path = self.ScanParser.run(appscanCMDexe, file_path, task_folder_path)
        
    #     if test:
    #         ## 測試
    #         if input("Generate Excel report? (Y/N): ").strip().upper() == "Y":
    #             self.ScanExcelGenerater.generate_excel_report(json_path, task_folder_path)
            
    #         if input("Generate Word report? (Y/N): ").strip().upper() == "Y":
    #             self.ScanWordGenerater.generate_word_report(json_path, word_title, task_folder_path)
    #     else:
    #         self.ScanExcelGenerater.generate_excel_report(json_path, task_folder_path)
    #         self.ScanWordGenerater.generate_word_report(json_path, word_title, task_folder_path)
        
    #     #刪除json檔
    #     # os.remove(json_path)
    
### exe使用
    def data_preprocessing(self, task_folder_path, appscanCMDexe="C:/Program Files (x86)/HCL/AppScan Standard/AppScanCMD.exe"):   # 預處理資料
        # appscan_parser
        self.ScanParser.run(task_folder_path, appscanCMDexe)
        # appscan_image
        self.ScanImageGenerater.generate_image(task_folder_path)
    
    def report_generater(self, data):    # 後產報告

        title_word = data["word_title"]
        report_type = data["report_type"]
        task_folder_path = data["task_folder_path"]

        self.ScanExcelGenerater.generate_excel_report(task_folder_path, report_type)
        self.ScanWordGenerater.generate_word_report(title_word, task_folder_path, report_type)
        # 刪除json檔
        # os.remove(json_path[0])

### api使用
    def run(self):
        appscanCMDexe = "C:/Program Files (x86)/HCL/AppScan Standard/AppScanCMD.exe"
        with self.Session() as session:
            # select all prepare
            tasks = session.query(tasklist).filter_by(status='prepare').all()
            comparetasks = session.query(comparelist).filter_by(status='prepare').all()
            for task in tasks:
                task.status='processing'

            for task in comparetasks:
                task.status='processing'
            
            session.commit()

            for task in tasks:
                try:
                    ## 確認appscan檔有沒有存在
                    project_path = os.path.join(os.getcwd(), "data/uploadproject", task.taskname, "appscan")
                    if not os.path.isdir(project_path):
                        continue                    
                    self.data_preprocessing(project_path, appscanCMDexe)
                    task.status='ok'
                except Exception as e:
                    print(f"An error occurred: {e}")
                    task.status='error'               
                session.commit()


#%%
# def _file_copy(file_path, task_folder_path):
#         new_file_path = []
#         if isinstance(task_folder_path, str):
#             task_folder_path = [task_folder_path] * len(file_path)

#         for ff, tt in zip(file_path, task_folder_path):
#             new_folder = dst_path = shutil.copy(ff, os.path.join(tt, os.path.basename(ff)))
#             new_file_path.append(new_folder)

#         return new_file_path


def main():
    report = Web_Report()
    report.run()

if __name__ == "__main__":
    main()
