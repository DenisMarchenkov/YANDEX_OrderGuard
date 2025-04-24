import requests
from logging_config import *

def get_orders(campaign_id, api_token, limit=20, page_token=None, offer_ids=None):
    url = f"https://api.partner.market.yandex.ru/campaigns/{campaign_id}/orders"
    headers = {
        'Api-Key': api_token,
        'Accept': 'application/json',
        'X-Market-Integration': 'OrderGuardTest'
    }
    params = {
        "limit": limit,
        "fake": 'false',
        "status": "PROCESSING",
        "substatus": "STARTED",
    }
    if page_token:
        params["page_token"] = page_token
    if offer_ids:
        params["offerId"] = ",".join(offer_ids) if isinstance(offer_ids, list) else offer_ids

    all_orders = []
    while True:
        try:
            response = requests.get(url, headers=headers, params=params)
            if response.status_code != 200:
                logging.error(f"Ошибка при запросе: {response.status_code}")
                logging.error(response.text)
                break
            data = response.json()
            orders = data.get("orders", [])
            all_orders.extend(orders)
            next_page_token = data.get("paging", {}).get("nextPageToken")
            if not next_page_token:
                break
            params["page_token"] = next_page_token
        except Exception as e:
            logging.exception(f"Ошибка при получении заказов: {e}")
            break
    return all_orders