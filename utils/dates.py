from datetime import datetime

def parse_date(date_str, fmt):
    try:
        return datetime.strptime(date_str, fmt).date()
    except Exception:
        return None