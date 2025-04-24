import smtplib
from logging_config import *
from platform import python_version
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from settings import SERVER_SMTP, LOGIN, PASSWORD_EMAIL_API, STORE, CONFIRMATION_REFUSED


def send_email(body, host, sender_email, sender_password, recipients, msg_subject, attachments=None):
    sender = sender_email
    msg = MIMEMultipart()
    msg['From'] = f'OrderGuard <{sender}>'
    msg['To'] = ', '.join([f'<{r}>' for r in recipients])
    msg['Subject'] = msg_subject
    msg['Reply-To'] = sender
    msg['Return-Path'] = sender
    msg['X-Mailer'] = 'Python/' + python_version()

    if attachments:
        for path in attachments:
            if path:
                try:
                    with open(path, "rb") as file:
                        part = MIMEBase('application', 'octet-stream')
                        part.set_payload(file.read())
                    encoders.encode_base64(part)
                    part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(path)}")
                    msg.attach(part)
                    logging.info(f"Файл {os.path.basename(path)} прикреплён к письму.")
                except Exception as e:
                    logging.warning(f"Ошибка при прикреплении {path}: {e}")

    msg.attach(MIMEText(body, 'plain'))

    try:
        server = smtplib.SMTP_SSL(host)
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipients, msg.as_string())
        server.quit()
        logging.info(f"Письмо с темой '{msg_subject}' отправлено успешно.")
        return True
    except Exception as e:
        logging.warning(f"Ошибка при отправке письма с темой '{msg_subject}': {e}")
        return False


def send_recap_email(report_file_path, labels_file_path, orders_ids):
    recipients = ["denismarchenkov@dfarm.ru"]  # укажи нужные email-адреса
    attachments = [f for f in [report_file_path, labels_file_path] if f]
    logging.info(f"Отправка email с вложениями: {[f for f in [report_file_path, labels_file_path] if f]}")
    subject = f"Подтверждение заказов YANDEX - {STORE}"
    body = (
        f"Уважаемые коллеги,\n\n"
        f"Во вложении направляем документы по заказам YANDEX – {STORE}:\n"
        f"– файл с наклейками для маркировки коробок,\n"
        f"– файл «рекап» с данными для комплектации заказов.\n\n"
        f"Номера заказов:\n"
    )

    body += "\n".join(f"- Заказ № {oid}" for oid in orders_ids)
    body += (
        "\n\nПросим использовать указанные материалы при сборке и отгрузке заказов.\n\n"
        "Если возникнут вопросы, пожалуйста, свяжитесь с ответственным сотрудником.\n\n"
        "С уважением,\n"
        "Автоматизированная система OrderGuard"
    )

    return send_email(
        body=body,
        host=SERVER_SMTP,
        sender_email=LOGIN,
        sender_password=PASSWORD_EMAIL_API,
        recipients=recipients,
        msg_subject=subject,
        attachments=attachments
    )


def send_refused_email(order_ids: list[str]):
    if not order_ids:
        return

    subject = f"Заказы YANDEX - {STORE} с отклонёнными позициями"
    body = (
        f"Уважаемые коллеги,\n\n"
        f"Сообщаем, что в следующих заказах YANDEX – {STORE} были выявлены отклонённые позиции (REFUSED ≠ 0):\n\n"
    )
    body += "\n".join(f"- Заказ № {oid}" for oid in order_ids)
    body += (
        "\n\nПросим проверить данные заказы и при необходимости предпринять соответствующие действия.\n\n"
        "Если возникнут вопросы, пожалуйста, свяжитесь с ответственным сотрудником.\n\n"
        "С уважением,\n"
        "Автоматизированная система OrderGuard"
    )

    attachments = [os.path.join(CONFIRMATION_REFUSED, f"{oid}.xls") for oid in order_ids]
    recipients = ["denismarchenkov@dfarm.ru"]  # замените на реальные email'ы

    send_email(
        body=body,
        host=SERVER_SMTP,
        sender_email=LOGIN,
        sender_password=PASSWORD_EMAIL_API,
        recipients=recipients,
        msg_subject=subject,
        attachments=attachments
    )