from logging_config import *
from settings import *

from utils.ftp_utils import download_from_ftp
from utils.report_utils import process_confirmation_files, handle_report
from utils.yandex_api.label_utils import handle_labels
from utils.email_utils import send_recap_email, send_refused_email
from utils.yandex_api.status_utils import handle_status_update


def main():
    try:
        # 1. Скачиваем подтверждения с FTP
        download_from_ftp(SERVER, USER_NAME, PASSWORD, CONFIRMATION_DIR, DIVISION_ID)

        # 2. Обрабатываем и объединяем данные + получаем заказы с отказом
        confirmations_df, refused_order_ids = process_confirmation_files(CONFIRMATION_DIR)

        if refused_order_ids:
            logging.info(f"Заказы с отклонёнными позициями (REFUSED != 0): {refused_order_ids}")
            send_refused_email(refused_order_ids)

        if not confirmations_df.empty:
            # 3. Отчёт
            report_file_path = handle_report(confirmations_df, RECAPS_DIR)

            # 4. Наклейки
            labels_file_path = handle_labels(confirmations_df, BUSINESS_ID, HEADERS, STICKERS_DIR)

            # 5. Статусы
            handle_status_update(confirmations_df, CAMPAIGN_ID, HEADERS)

            # 6. Отправка отчета и наклеек по почте
            if report_file_path or labels_file_path:
                send_recap_email(
                    report_file_path=report_file_path,
                    labels_file_path=labels_file_path,
                    orders_ids=confirmations_df["HDRTAG2"].to_list()
                )
            else:
                logging.info("Нет файлов для отправки по почте.")
        else:
            logging.info("Нет данных для создания отчета, получения наклеек и обновления статусов.")

    except Exception as e:
        logging.exception(f"Ошибка в процессе выполнения main(): {e}")


if __name__ == "__main__":
    main()
