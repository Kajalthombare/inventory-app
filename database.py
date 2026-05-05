import os
from sqlalchemy import create_engine

DATABASE_URL = os.getenv("MYSQL_URL")

engine = create_engine(
    DATABASE_URL,
    pool_pre_ping=True
)