import ftplib
import time
from logging_config import *


def download_from_ftp(server, username, password, local_directory, division_id, max_retries=3, retry_delay=5):
    remote_directory = "/OZON_orders/Frenchpharmacy/Confirmations"
    attempt = 0

    while attempt < max_retries:
        try:
            with ftplib.FTP(server) as ftp:
                ftp.login(user=username, passwd=password)
                logging.info("Успешное подключение и авторизация на сервере")
                ftp.cwd(remote_directory)
                logging.info("Успешный переход в директорию на сервере")

                if not os.path.exists(local_directory):
                    os.makedirs(local_directory)

                files = ftp.nlst()
                found_any = False

                for filename in files:
                    if (
                        filename.endswith('.xls') and
                        '_' in filename and
                        filename.rsplit('_', 1)[-1].replace('.xls', '') == division_id
                    ):
                        found_any = True
                        order_number = filename.rsplit('_', 1)[0]
                        new_filename = f"{order_number}.xls"
                        local_path = os.path.join(local_directory, new_filename)

                        with open(local_path, 'wb') as local_file:
                            ftp.retrbinary(f"RETR {filename}", local_file.write)
                            logging.info(f"Файл скачан и переименован: {filename} → {new_filename}")

                        ftp.delete(filename)
                        logging.info(f"Файл удален с сервера: {filename}")

                if not found_any:
                    logging.info(f"Файлы с кодом подразделения {division_id} не найдены в директории.")

                return

        except (ftplib.all_errors, ConnectionError) as e:
            logging.error(f"Ошибка при работе с FTP: {e}. Попытка {attempt + 1} из {max_retries}")
            time.sleep(retry_delay)
            attempt += 1

    logging.error("Не удалось подключиться к FTP-серверу после нескольких попыток")



def upload_file_to_ftp(server, username, password, store, local_path_file):
    remote_directory = f'/YANDEX_orders/{store}/Orders/'
    try:
        with ftplib.FTP(server) as ftp:
            try:
                ftp.login(user=username, passwd=password)
                logging.info("Успешное подключение и авторизация на сервере")
            except ftplib.error_perm as e:
                logging.error(f"Ошибка авторизации: {e}")
                return

            filename = os.path.basename(local_path_file)

            def upload_to_directory(directory):
                try:
                    ftp.cwd(directory)
                    logging.info(f"Успешный переход в директорию {directory}")
                except ftplib.error_perm:
                    logging.warning(f"Директория {directory} не найдена, создаем...")
                    try:
                        ftp.mkd(directory)
                        ftp.cwd(directory)
                        logging.info(f"Директория {directory} создана и выбрана")
                    except ftplib.error_perm as e:
                        logging.error(f"Ошибка при создании директории {directory}: {e}")
                        return
                try:
                    with open(local_path_file, 'rb') as file:
                        ftp.storbinary(f'STOR {filename}', file)
                        logging.info(f"Файл загружен: {filename} в {directory}")
                except ftplib.all_errors as e:
                    logging.error(f"Ошибка при загрузке файла {filename} в {directory}: {e}")

            upload_to_directory(remote_directory)
    except ftplib.all_errors as e:
        logging.error(f"Ошибка соединения с FTP-сервером: {e}")