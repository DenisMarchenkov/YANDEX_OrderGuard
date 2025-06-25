import shutil
import pandas as pd
from datetime import datetime
from settings import CONFIRMATION_PROCESSED, CONFIRMATION_REFUSED, STORE_SUFFIX
from utils.excel_formatting import format_report_sheets
from logging_config import *


def append_data_from_file(file_path, df=None):
    try:
        file_data = pd.read_excel(file_path, dtype={'CODEART': str, 'CODEPST': str})
        df = pd.concat([df, file_data], ignore_index=True) if df is not None else file_data
        return df.fillna('нет данных')
    except Exception as e:
        logging.error(f"Ошибка при чтении файла {file_path}: {e}")
        return df


def process_confirmation_files(source_dir, processed_dir=CONFIRMATION_PROCESSED, refused_dir=CONFIRMATION_REFUSED):
    all_dataframes = []
    refused_order_ids = set()

    os.makedirs(processed_dir, exist_ok=True)
    os.makedirs(refused_dir, exist_ok=True)

    for file in os.listdir(source_dir):
        filepath = os.path.join(source_dir, file)

        if not os.path.isfile(filepath) or not file.endswith(".xls"):
            continue

        processed_path = os.path.join(processed_dir, file)
        refused_path = os.path.join(refused_dir, file)

        # Проверка: файл уже есть в обработанных или в отказах
        if os.path.exists(processed_path):
            logging.info(f"Файл уже обработан ранее: {file}")
            continue

        try:
            df = append_data_from_file(filepath)
            if df is not None:
                if 'REFUSED' in df.columns and 'HDRTAG2' in df.columns:
                    refused_ids = set(df[df['REFUSED'] != 0]['HDRTAG2'].unique())
                    if refused_ids:
                        refused_order_ids.update(refused_ids)
                        filtered_df = df[~df['HDRTAG2'].isin(refused_ids)]
                        if not filtered_df.empty:
                            all_dataframes.append(filtered_df)
                        shutil.move(str(filepath), str(refused_path))
                        logging.info(f"Файл с отказами перемещён в: {refused_path}")
                        continue  # пропускаем обычное перемещение
                    else:
                        all_dataframes.append(df)
                else:
                    all_dataframes.append(df)

            shutil.move(str(filepath), str(processed_path))
            logging.info(f"Файл перемещён в обработанные: {processed_path}")
        except Exception as e:
            logging.error(f"Ошибка при обработке файла {file}: {e}")

    combined_df = pd.concat(all_dataframes, ignore_index=True) if all_dataframes else pd.DataFrame()
    return combined_df, list(refused_order_ids)



def save_orders_to_excel(dataframe, output_path):
    if not os.path.exists(output_path):
        os.makedirs(output_path)

    id_recap = datetime.now().strftime("%d%m%y-%H%M%S")
    file_name = f'recap-YANDEX-{id_recap}.xlsx'
    path_file = os.path.join(output_path, file_name)

    with pd.ExcelWriter(path_file) as writer:
        all_orders = dataframe.pivot_table(index=['HDRTAG2', 'HDRTAG1'], values='QNT', aggfunc='sum').reset_index()
        all_orders.to_excel(writer, sheet_name='Список заказов', index=False)

        pivot_table = dataframe.pivot_table(index=['FIRM', 'CODEART', 'NAME', 'GDATE'], values='QNT', aggfunc='sum').reset_index()
        pivot_table.to_excel(writer, sheet_name='Сводная таблица', index=False)

        dataframe.to_excel(writer, columns=['HDRTAG2', 'FIRM', 'CODEART', 'NAME', 'QNT'], sheet_name="Лист подбора", index=False)

        for order_id, order_data in dataframe.groupby('HDRTAG2'):
            order_data.drop(columns=['PODRCD', 'REFUSED', 'CODEPST'], errors='ignore')\
                .to_excel(writer, sheet_name=f'ORDER {order_id}', index=False)

    return path_file


def handle_report(df, recaps_dir):
    report_file_path = save_orders_to_excel(df, recaps_dir)
    format_report_sheets(report_file_path, STORE_SUFFIX)
    return report_file_path
