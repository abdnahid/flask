from pymongo import MongoClient
import os
from dotenv import load_dotenv


load_dotenv()


def connect_db():
    cluster = MongoClient(os.getenv("MONGO_URI"))
    db = cluster["dmiAuto"]
    return db
