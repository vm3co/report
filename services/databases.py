import os
from sqlalchemy import create_engine
from sqlalchemy import Column, Integer, String
from sqlalchemy.ext.declarative import declarative_base

# engine = create_engine('sqlite:////'+os.path.join(os.getcwd(),"data/vultask.db"))
engine = create_engine('sqlite:///'+os.path.join(os.getcwd(),"data/vultask.db"))
db = declarative_base()

class tasklist(db):
    __tablename__ = "tasklist"
    id = Column(Integer, primary_key=True)
    taskname = Column(String)
    exectime = Column(String)
    status = Column(String)
    ip = Column(Integer)
    critical = Column(Integer)
    high = Column(Integer)
    medium = Column(Integer)
    low = Column(Integer)
    info = Column(Integer)

    def __init__(self, taskname, exectime):
        self.taskname = taskname
        self.exectime = exectime
        self.status='pedding'
        self.ip = 0
        self.critical =0
        self.high = 0
        self.medium =0
        self.low = 0
        self.info = 0

class comparelist(db):
    __tablename__ = "comparelist"
    id = Column(Integer, primary_key=True)
    taskname = Column(String)
    exectime = Column(String)
    status = Column(String)
    ip = Column(Integer)
    critical = Column(Integer)
    high = Column(Integer)
    medium = Column(Integer)
    low = Column(Integer)
    info = Column(Integer)

    def __init__(self,taskname, exectime):
        self.taskname = taskname
        self.exectime = exectime
        self.status='pedding'
        self.ip = 0
        self.critical =0
        self.high = 0
        self.medium =0
        self.low = 0
        self.info = 0

class adminuser(db):
    __tablename__ = "adminuser"
    id = Column(Integer, primary_key=True)
    username = Column(String)
    password = Column(String) 

if __name__ == "__main__":
    # === 建立資料庫 ===
    db.metadata.create_all(engine)
