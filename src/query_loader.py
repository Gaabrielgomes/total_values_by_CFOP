from pathlib import Path


BASE_DIR = Path(__file__).parent / "queries"

def get_query(name):
    with open(BASE_DIR / f"{name}.sql", "r", encoding="utf-8") as f:
        return f.read()
