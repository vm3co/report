import os
from sqlalchemy.orm import sessionmaker
import pandas as pd

from services.nessus_parser import NessusParser
from services.nessus_excel import NessusExcelGenerater
from services.nessus_word import NessusDocGenerater
from services.nessus_image import NessusImageGenerater

from services.appscan_parser import AppScanParser
from services.appscan_excel import AppScanExcelReport
from services.appscan_word import AppScanWordReport
from services.appscan_image_generater import AppScanImageGenerater

from services.databases import tasklist, comparelist, adminuser, engine

class ReportMaker:
    def __init__(self):
        self.Session = sessionmaker(bind=engine)

        self.nessus_parser = NessusParser()
        self.nessus_excel_generater = NessusExcelGenerater()
        self.nessus_word_generater = NessusDocGenerater()
        self.nessus_image_generater = NessusImageGenerater()
        self.ScanParser = AppScanParser()
        self.ScanExcelGenerater = AppScanExcelReport()
        self.ScanWordGenerater = AppScanWordReport()
        self.ScanImageGenerater = AppScanImageGenerater()

    def report_generater(self, data):    # 後產報告
        title_word = data["word_title"]
        report_type = data["report_type"]
        task_folder_path = data["task_folder_path"]

        task_folder_web_path = os.path.join(task_folder_path, "appscan")
        task_folder_va_path = os.path.join(task_folder_path, "nessus")
        
        if os.path.isdir(task_folder_web_path):
            self.ScanExcelGenerater.generate_excel_report(task_folder_web_path, report_type)
            self.ScanWordGenerater.generate_word_report(title_word, task_folder_web_path, report_type)

        if os.path.isdir(task_folder_va_path):
            self.nessus_excel_generater.generate_report(task_folder_va_path)
            self.nessus_word_generater.generate_report(title_word, task_folder_va_path)

        # 刪除json檔
        # os.remove(json_path[0])


    def run(self):
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
                    project_nessus_path = os.path.join(os.getcwd(), "data/uploadproject", task.taskname, "nessus")
                    ## 確認nessus檔有沒有存在
                    if os.path.isdir(project_nessus_path):
                        df_nessus, df_nessus_scan_info = self.nessus_parser.nessus_to_df(project_nessus_path)

                        # image_data
                        df_count, host_count = self.nessus_image_generater.image_data(df_nessus)
                        self.nessus_image_generater.generate_image(df_count, project_nessus_path, "bar_summary_va.jpg")
                        self.nessus_image_generater.generate_image(host_count, project_nessus_path, "summary_host_va.jpg")
                        
                        task.ip = int(pd.to_numeric(df_nessus_scan_info.get("count", 0), errors="coerce").fillna(0).astype(int).sum())

                        severity_counts = df_nessus["severity"].value_counts().reindex(["4", "3", "2", "1", "0"], fill_value=0)
                        task.critical = int(severity_counts["4"])
                        task.high = int(severity_counts["3"])
                        task.medium = int(severity_counts["2"])
                        task.low = int(severity_counts["1"])
                        task.info = int(severity_counts["0"])                    
                    ## 確認appscan檔有沒有存在
                    project_appscan_path = os.path.join(os.getcwd(), "data/uploadproject", task.taskname, "appscan")
                    if os.path.isdir(project_appscan_path): 
                        appscanCMDexe = "C:/Program Files (x86)/HCL/AppScan Standard/AppScanCMD.exe"                  
                        # appscan_parser
                        self.ScanParser.run(project_appscan_path, appscanCMDexe)
                        # appscan_image
                        self.ScanImageGenerater.generate_image(project_appscan_path)

                    task.status='ok'

                except Exception as e:
                    print(f"An error occurred: {e}")
                    task.status='error'
                
                session.commit()
            #comparetasks = session.query(comparelist).filter_by(status='pedding').all()

            # for task in comparetasks:
            #     try:
            #         scantask = task.taskname.split('/')
            #         firstfilepath = os.path.join(os.getcwd(), "data/uploadproject", scantask[0])
            #         secondfilepath = os.path.join(os.getcwd(), "data/uploadproject", scantask[1])
            #         targetfilepath = os.path.join(os.getcwd(), "data/compareproject", f"{scantask[0]}_{scantask[1]}")

            #         countip,count_severity = self.generatecompareexcel.analysisnessuscompare(firstfilepath, secondfilepath, targetfilepath)
            #         self.generateimage.generateimage(targetfilepath)
            #         self.generatejsondata.generatejsondata(targetfilepath)
            #         self.generatedocreport.generatedocreport(targetfilepath)

            #         task.ip = int(countip)
            #         task.critical = int(count_severity['4'])
            #         task.high = int(count_severity['3'])
            #         task.medium = int(count_severity['2'])
            #         task.low = int(count_severity['1'])
            #         task.info = int(count_severity['0'])
            #         task.status='ok'
                
            #     except Exception as e:
            #         print(f"An error occurred: {e}")
            #         task.status='error'
                
            #     session.commit()

report_maker = ReportMaker()

def main():
    report_maker.run()

if __name__ == "__main__":
    main()