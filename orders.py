from logging_config import *
from settings import CAMPAIGN_ID, API_TOKEN, CUSTOMER_ID_IN_SUPPLIER_CRM, DIVISION_ID, ORDERS_DIR
from settings import SERVER, USER_NAME, PASSWORD, STORE

from utils.yandex_api.orders_utils import get_orders
from utils.dbf_utils import export_orders_to_dbf_files


def main():
    try:
        logging.info("Запуск получения заказов с Яндекс Маркета")
        ftp_config = {
            "server": SERVER,
            "username": USER_NAME,
            "password": PASSWORD,
            "store": STORE
        }

        # 1. Получаем данные о заказах
        orders = get_orders(CAMPAIGN_ID, API_TOKEN)
        logging.info(f"Получено заказов: {len(orders)}")

        # 2. Экспортируем полученные данные в dbf файлы, выгружаем на ftp
        export_orders_to_dbf_files(
            orders=orders,
            output_dir=ORDERS_DIR,
            customer_id=CUSTOMER_ID_IN_SUPPLIER_CRM,
            division_id=DIVISION_ID,
            ftp_config=ftp_config
        )

        logging.info("Завершение обработки заказов")

    except Exception as e:
        logging.exception(f"Ошибка в процессе выполнения main(): {e}")

if __name__ == "__main__":
    main()
