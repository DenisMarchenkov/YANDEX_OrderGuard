import time
import requests
from logging_config import *


def validate_dataframe(df):
    if "HDRTAG2" not in df.columns:
        logging.error("Колонка 'HDRTAG2' не найдена в DataFrame.")
        return False
    return True


def extract_unique_orders(df):
    try:
        return df.drop_duplicates(subset=["HDRTAG2"])
    except Exception as e:
        logging.error(f"Ошибка при удалении дубликатов: {e}")
        return None


def build_order_payload(df, new_status, new_substatus):
    orders = []
    ids_sent = []
    for _, row in df.iterrows():
        try:
            order_id = int(row["HDRTAG2"])
            orders.append({
                "id": order_id,
                "status": new_status,
                "substatus": new_substatus
            })
            ids_sent.append(order_id)
        except Exception as e:
            logging.warning(f"Ошибка при обработке строки HDRTAG2={row.get('HDRTAG2')}: {e}")
    return orders, ids_sent


def send_status_update_request(payload, campaign_id, headers, max_retries=3, retry_delay=3):
    url = f"https://api.partner.market.yandex.ru/campaigns/{campaign_id}/orders/status-update"
    logging.info(f"Отправка {len(payload['orders'])} заказов на обновление статусов...")

    for attempt in range(1, max_retries + 1):
        try:
            response = requests.post(url, headers=headers, json=payload)
            if response.status_code == 200:
                logging.info("Статусы заказов успешно обновлены.")
                return True
            elif response.status_code in (429, 500, 502, 503, 504):
                logging.warning(f"Временная ошибка {response.status_code}, попытка {attempt}/{max_retries}")
                if attempt < max_retries:
                    time.sleep(retry_delay * attempt)
                    continue
            else:
                logging.error(f"Ошибка {response.status_code}: {response.text}")
                return False
        except requests.RequestException as e:
            logging.error(f"Сетевая ошибка, попытка {attempt}/{max_retries}: {e}")
            if attempt < max_retries:
                time.sleep(retry_delay * attempt)
            else:
                return False
        except Exception as e:
            logging.exception(f"Непредвиденная ошибка: {e}")
            return False
    return False


def update_order_statuses(df, campaign_id, headers, max_retries=3, retry_delay=3):
    if not validate_dataframe(df):
        return None

    unique_df = extract_unique_orders(df)
    if unique_df is None:
        return None

    orders, ids_sent = build_order_payload(unique_df, "PROCESSING", "READY_TO_SHIP")
    if not orders:
        logging.warning("Нет заказов для отправки.")
        return []

    payload = {"orders": orders}
    success = send_status_update_request(payload, campaign_id, headers, max_retries, retry_delay)
    return ids_sent if success else None


def handle_status_update(df, campaign_id, headers):
    updated = update_order_statuses(df, campaign_id, headers)
    if updated:
        logging.info(f"Обновлены заказы: {updated}")
    else:
        logging.info("Статусы заказов не обновлены или не было заказов для обновления.")
