import os
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, declarative_base

# ✅ GET FROM RAILWAY ENV
DATABASE_URL = os.getenv("MYSQL_URL")

if not DATABASE_URL:
    raise ValueError("MYSQL_URL not found in environment")

# ✅ FIX MYSQL URL FORMAT (IMPORTANT)
DATABASE_URL = DATABASE_URL.replace("mysql://", "mysql+pymysql://")

engine = create_engine(DATABASE_URL, pool_pre_ping=True)

SessionLocal = sessionmaker(bind=engine)

Base = declarative_base()