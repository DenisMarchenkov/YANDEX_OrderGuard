import dbf

from logging_config import *
from utils.dates import parse_date
from utils.ftp_utils import upload_file_to_ftp

def export_orders_to_dbf_files(orders, output_dir, customer_id, division_id, ftp_config):
    os.makedirs(output_dir, exist_ok=True)
    for order in orders:
        order_id = order.get("id")
        order_date_str = order.get("creationDate")
        shipment_data = order.get("delivery", {}).get("shipments", [])
        items = order.get("items", [])
        if not items:
            continue

        order_date = parse_date(order_date_str, "%d-%m-%Y %H:%M:%S")
        shipment_date = parse_date(shipment_data[0].get("shipmentDate"), "%d-%m-%Y") if shipment_data else None

        filename = os.path.join(output_dir, f"{order_id}.dbf")
        if os.path.exists(filename):
            logging.info(f"Файл уже существует, пропускаю: {filename}")
            continue

        table = dbf.Table(
            filename,
            'offer_id C(50); name C(255); price N(10,2); qty N(10,0); post_num C(50); '
            'ord_date D; ship_date D; comment C(255); cust_id C(10); div_id C(10)',
            codepage='cp866'
        )
        table.open(mode=dbf.READ_WRITE)
        for item in items:
            table.append((
                item.get("offerId", ""),
                item.get("offerName", ""),
                float(item.get("price", 0)),
                int(item.get("count", 0)),
                str(order_id),
                order_date,
                shipment_date,
                f'{order_id} Заказ YANDEX Frenchpharmacy',
                str(customer_id),
                str(division_id)
            ))
        table.close()
        logging.info(f"Создан файл: {filename}")
        upload_file_to_ftp(
            server=ftp_config["server"],
            username=ftp_config["username"],
            password=ftp_config["password"],
            store=ftp_config["store"],
            local_path_file=filename
        )