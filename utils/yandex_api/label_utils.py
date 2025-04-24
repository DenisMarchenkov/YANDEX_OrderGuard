import time
import requests
from logging_config import *


def request_report_generation(order_ids, business_id, headers, format_pdf):
    url = "https://api.partner.market.yandex.ru/reports/documents/labels/generate"
    params = {'format': format_pdf}
    payload = {
        "businessId": business_id,
        "orderIds": order_ids,
        "sortingType": "SORT_BY_GIVEN_ORDER"
    }
    try:
        response = requests.post(url, headers=headers, params=params, json=payload)
        response.raise_for_status()
        logging.info("Запрос на генерацию ярлыков отправлен успешно.")
        return response.json().get('result', {}).get('reportId')
    except requests.RequestException as e:
        logging.error(f"Ошибка при генерации отчета: {e}")
        return None


def poll_report_status(report_id, headers, max_retries=5, retry_delay=5):
    url = f"https://api.partner.market.yandex.ru/reports/info/{report_id}"
    for attempt in range(max_retries):
        logging.info(f"Проверка статуса ярлыков (попытка {attempt + 1}/{max_retries})...")
        try:
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            result = response.json().get('result', {})
            status = result.get('status')
            if status == 'DONE':
                return result.get('file')
            elif status in ['PROCESSING', 'NEW', 'PENDING']:
                time.sleep(retry_delay)
            else:
                logging.error(f"Неподдерживаемый статус: {status}, детали: {result.get('error')}")
                return None
        except requests.RequestException as e:
            logging.error(f"Ошибка при проверке статуса отчета: {e}")
            return None
    logging.error("Превышено количество попыток проверки статуса ярлыков.")
    return None


def download_file(file_url, save_path, report_id, timeout=15, retries=3, delay=5):
    filename = f"labels_{report_id}.pdf"
    file_path = os.path.join(save_path, filename)
    os.makedirs(save_path, exist_ok=True)

    for attempt in range(1, retries + 1):
        try:
            logging.info(f"Скачивание файла {attempt}/{retries}: {file_url}")
            with requests.get(file_url, stream=True, timeout=timeout) as r:
                r.raise_for_status()
                with open(file_path, 'wb') as f:
                    for chunk in r.iter_content(chunk_size=8192):
                        f.write(chunk)
            logging.info(f"Файл ярлыков сохранён: {file_path}")
            return file_path
        except requests.RequestException as e:
            logging.warning(f"Ошибка при скачивании ярлыков, попытка {attempt}: {e}")
            if attempt < retries:
                time.sleep(delay)

    logging.error(f"Не удалось скачать файл ярлыков после {retries} попыток.")
    return None


def generate_order_labels(order_ids, business_id, headers, format_pdf='A7',
                          save_path=None, retry_delay=5, download_timeout=15):
    report_id = request_report_generation(order_ids, business_id, headers, format_pdf)
    if not report_id:
        return None

    file_url = poll_report_status(report_id, headers, max_retries=5, retry_delay=retry_delay)
    if not file_url:
        return None

    return download_file(file_url, save_path, report_id, timeout=download_timeout) if save_path else file_url


def handle_labels(df, business_id, headers, save_path):
    order_ids = df['HDRTAG2'].unique().tolist()
    return generate_order_labels(order_ids, business_id, headers, save_path=save_path)
