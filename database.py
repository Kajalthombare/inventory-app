import os
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, declarative_base

# Load local .env file if it exists
if os.path.exists(".env"):
    with open(".env", "r", encoding="utf-8") as f:
        for line in f:
            line_str = line.strip()
            if "=" in line_str and not line_str.startswith("#"):
                key, val = line_str.split("=", 1)
                key = key.strip()
                val = val.strip().strip("'\"")
                if key and val and key not in os.environ:
                    os.environ[key] = val

# Comprehensive detection of database URL across cloud providers (Railway, Render, Supabase, Neon, Docker)
DATABASE_URL = (
    os.environ.get("DATABASE_URL") or
    os.environ.get("POSTGRES_URL") or
    os.environ.get("POSTGRESQL_URL") or
    os.environ.get("DATABASE_PUBLIC_URL") or
    os.environ.get("DATABASE_PRIVATE_URL") or
    os.environ.get("MYSQL_URL")
)

# Also check individual PG* variables if URL is not explicitly set
if not DATABASE_URL and os.environ.get("PGHOST") and os.environ.get("PGDATABASE"):
    user = os.environ.get("PGUSER", "postgres")
    password = os.environ.get("PGPASSWORD", "")
    host = os.environ.get("PGHOST")
    port = os.environ.get("PGPORT", "5432")
    dbname = os.environ.get("PGDATABASE")
    if password:
        DATABASE_URL = f"postgresql://{user}:{password}@{host}:{port}/{dbname}"
    else:
        DATABASE_URL = f"postgresql://{user}@{host}:{port}/{dbname}"

# Enforce presence of DATABASE_URL – no silent SQLite fallback
if not DATABASE_URL:
    raise RuntimeError(
        "❗ DATABASE_URL not set – the app must connect to PostgreSQL. "
        "Check Railway > Settings > Variables."
    )

# Dynamic driver selection for PostgreSQL
if DATABASE_URL.startswith("postgres://") or (DATABASE_URL.startswith("postgresql://") and not DATABASE_URL.startswith("postgresql+")):
    pg_driver = "postgresql+psycopg2://"
    try:
        import psycopg2
    except ImportError:
        try:
            import psycopg
            pg_driver = "postgresql+psycopg://"
        except ImportError:
            pg_driver = "postgresql+psycopg2://"
    
    if DATABASE_URL.startswith("postgres://"):
        DATABASE_URL = DATABASE_URL.replace("postgres://", pg_driver, 1)
    else:
        DATABASE_URL = DATABASE_URL.replace("postgresql://", pg_driver, 1)

# MySQL driver fix
if DATABASE_URL.startswith("mysql://"):
    DATABASE_URL = DATABASE_URL.replace("mysql://", "mysql+pymysql://", 1)

# Engine initialization
if DATABASE_URL.startswith("sqlite"):
    engine = create_engine(DATABASE_URL, connect_args={"check_same_thread": False})
else:
    engine = create_engine(DATABASE_URL, pool_pre_ping=True, pool_recycle=300)

SessionLocal = sessionmaker(bind=engine)
Base = declarative_base()

# Safe masked log for diagnostics
_masked = DATABASE_URL
if "@" in _masked:
    _parts = _masked.split("@")
    _cred = _parts[0].split(":")
    if len(_cred) > 2:
        _masked = f"{_cred[0]}:{_cred[1]}:****@{_parts[1]}"
print(f"📦 Database configured: {_masked}")
