import xml.etree.ElementTree as ET
import os, json
import pandas as pd
from datetime import datetime, UTC

class NessusParser:
    def __init__(self):
        print("NessusParser --- loaded")

    @staticmethod
    def output(output):
        lines = [line.strip() for line in output.splitlines() if line.strip()]
        plugin_output = ""
        for line in lines:
            plugin_output = plugin_output + line.strip() + "\n"
        return plugin_output.strip()
        
    def nessus_to_df(self, taskpath):
        """ .nessus to DataFrame"""
        nessus_files = []
        for root, _, files in os.walk(taskpath):
            nessus_files.extend([os.path.join(root, f) for f in files if f.endswith('.nessus')])
        
        data = []
        scan_info = []
        
        for file_no, file in enumerate(nessus_files):
            tree = ET.parse(file)
            root = tree.getroot()

            # 初始化
            scan_details = {}
            target_value = []
            start_time, end_time = "Unknown", "Unknown"
            
            # 計算不重複的 ReportHost 數量
            unique_hosts = {host.get("name") for host in root.findall(".//ReportHost")}
            count = len(unique_hosts)  

            for pref in root.findall(".//preference"):
                name = pref.find("name").text if pref.find("name") is not None else ""
                value = pref.find("value").text if pref.find("value") is not None else ""

                if name == "TARGET":
                    target_value = value

                elif name == "scan_start_timestamp":
                    start_time = datetime.fromtimestamp(int(value), tz=UTC).strftime("%Y.%m.%d %H:%M:%S")
                elif name == "scan_end_timestamp":
                    end_time = datetime.fromtimestamp(int(value), tz=UTC).strftime("%Y.%m.%d %H:%M:%S")

            scan_details = {
                "file_no": file_no,
                "start": f"{start_time}",
                "end": f"{end_time}",
                "target": f"{target_value}",
                "count": f"{count}"
            }
            
            for host in root.findall(".//ReportHost"):
                host_ip = host.get("name")
                os_name, os_system = "", ""

                for prop in host.findall("HostProperties/tag"):
                    if prop.get("name") == "os":
                        os_name = prop.text
                    elif prop.get("name") == "operating-system":
                        os_system = prop.text

                for item in host.findall("ReportItem"):
                    plugin_id = item.attrib.get("pluginID")
                    severity = item.attrib.get("severity", "0")

                    if plugin_id == "45410" or int(severity) > 0:
                        plugin_id = item.attrib.get("pluginID")
                        plugin_name = item.attrib.get("pluginName")
                        description = item.findtext("description", "")
                        solution = item.findtext("solution", "")
                        see_also = item.findtext("see_also", "")
                        plugin_output = self.output(item.findtext("plugin_output", ""))

                        port = item.attrib.get("port")
                        protocol = item.attrib.get("protocol")

                        cve = item.findtext("cve", "")
                        cvss3 = item.findtext("cvss3_base_score", "")
                        cvss2 = item.findtext("cvss_base_score", "")
                        
                        try:
                            host_name = plugin_output.split("The host name known by Nessus is :")[1].split("The Common Name in the certificate is :")[0]
                        except:
                            host_name = ""

                        data.append([
                            file_no, host_ip, os_system, os_name, plugin_id, plugin_name, host_name, 
                            port, protocol, severity, description, solution, cve, cvss3, cvss2, see_also, plugin_output
                        ])
                    
            scan_info.append(scan_details)
        
        df_vulnerabilities = pd.DataFrame(data, columns=[
            "file_no", "ip", "system", "os", "pluginid", "pluginname", "hostname", "port", "protocol",
            "severity", "description", "solution", "cve", "cvss3", "cvss2", "see_also", "plugin_output"
        ])
        
        df_scan_info = pd.DataFrame(scan_info, columns=["file_no", "start", "end", "target", "count"])
        
        information = {
            "df_nessus": df_vulnerabilities, 
            "df_nessus_scan_info": df_scan_info
            }
        
        json_data = {key: df.to_dict(orient="records") for key, df in information.items()}                        
        # 存成 JSON 檔案
        with open(os.path.join(taskpath, "nessus.json"), "w", encoding="utf-8") as f:
            json.dump(json_data, f, indent=4, ensure_ascii=False)
        
        return df_vulnerabilities, df_scan_info


def main():
    taskpath = r"D:\GitHub\appliance\app\data\uploadproject\TEST3"
    temp = NessusParser()
    df_nessus, df_nessus_scan_info = temp.nessus_to_df(taskpath)
    print("df")
    return df_nessus, df_nessus_scan_info

if __name__ == "__main__":
    df_nessus, df_nessus_scan_info = main()