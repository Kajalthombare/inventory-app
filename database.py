# Database setup - SQLite (persistent via Railway volume)
import os
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, declarative_base

# Load local .env file if it exists
if os.path.exists(".env"):
    with open(".env") as f:
        for line in f:
            if "=" in line and not line.strip().startswith("#"):
                key, val = line.strip().split("=", 1)
                os.environ[key.strip()] = val.strip()

# Use SQLite — stored at /data/inventory.db on Railway (persistent volume)
# Falls back to local inventory.db for development
DATABASE_URL = os.environ.get("DATABASE_URL", "sqlite:////data/inventory.db")

# SQLite needs check_same_thread=False for FastAPI
engine = create_engine(DATABASE_URL, connect_args={"check_same_thread": False})

SessionLocal = sessionmaker(bind=engine)
Base = declarative_base()