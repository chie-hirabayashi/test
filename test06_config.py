import os
from dotenv import load_dotenv
from playhouse.db_url import connect

# import logging


print(os.environ.get("PWD"))

load_dotenv()

db = os.environ.get("DATABASE")

if not db.connect():
    print("接続NG")
    exit()
