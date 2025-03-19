import os
import time

import uno
import urllib.parse

class MenuUpdater:
    def __init__(self, host="localhost", port=2002):
        self.host = host
        self.port = port
        self.context = self._connect_to_libreoffice()
        self.desktop = self.context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", self.context)

    def _connect_to_libreoffice(self):
        local_context = uno.getComponentContext()
        resolver = local_context.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_context
        )
        return resolver.resolve(f"uno:socket,host={self.host},port={self.port};urp;StarOffice.ComponentContext")

    def re_index(self, doc_path):
        abs_path = os.path.abspath(doc_path)
        file_url = "file://" + urllib.parse.quote(abs_path)

        hidden_property = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        hidden_property.Name = "Hidden"
        hidden_property.Value = uno.Bool(False)
        props = (hidden_property,)
        
        doc = self.desktop.loadComponentFromURL(file_url, "_blank", 0, props)

        try:
            indexes = doc.getDocumentIndexes()
            for i in range(indexes.getCount()):
                index = indexes.getByIndex(i)
                index.update()

            time.sleep(2)
            doc.store()
        
        except Exception as e:
            print(f"{e}")
        
        finally:
            doc.close(True)

# 使用範例
if __name__ == "__main__":
    updater = MenuUpdater()
    updater.update_index("outputfile.docx")
