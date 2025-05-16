# -*- coding: utf-8 -*-
import imaplib
import email
import os
import re
import json
from email.header import decode_header
from datetime import date, datetime


IMAP_SERVER = 'imap.mail.ru'
# ЗАМЕНИТЕ НА ВАШИ ДАННЫЕ
EMAIL_ACCOUNT = '' 
APP_PASSWORD = '' 
DOWNLOAD_FOLDER = 'downloaded_calls'
MAILBOX_FOLDER = 'Calls'
METADATA_FILENAME = 'metadata.json'
MIN_CALL_DURATION_SECONDS = 15 

def decode_filename(filename_header):
    """Декодирует имя файла из заголовка письма."""
    filename = ""
    if filename_header:
        filename_decoded, encoding = decode_header(filename_header)[0]
        if isinstance(filename_decoded, bytes):
            try: filename = filename_decoded.decode(encoding if encoding else 'utf-8', errors='replace')
            except LookupError: filename = filename_decoded.decode('utf-8', errors='replace')
        else: filename = filename_decoded
    return filename.strip()

def extract_id_from_filename(filename):
    """Извлекает ID (цифры после 'vpbx') из имени файла."""
    if not filename: return None
    match = re.search(r'vpbx(\d+)', filename, re.IGNORECASE)
    if match: return match.group(1)
    else:
        print(f"    Предупреждение: Не удалось извлечь ID из имени файла '{filename}'.")
        base_filename = os.path.splitext(filename)[0]
        invalid_chars = r'[:/\\?*"<>|\t\n\r\f\v]'
        sanitized_name = re.sub(invalid_chars, '_', base_filename)
        sanitized_name = re.sub(r'_+', '_', sanitized_name).strip('_')
        return sanitized_name if sanitized_name else "unknown_call"

def extract_metadata(body, call_id, call_date_str_folder, original_filename):
    """Формирует метаданные для JSON файла и извлекает длительность в секундах."""
    metadata = {}
    duration_seconds = -1
    print("  Формирование метаданных...")

    metadata['call_id'] = call_id
    metadata['call_date_folder'] = call_date_str_folder
    iso_formatted_date = None
    if original_filename:
        datetime_match = re.search(r'(\d{4})\.(\d{2})\.(\d{2})__(\d{2})-(\d{2})-(\d{2})', original_filename)
        if datetime_match:
            try:
                year, month, day, hour, minute, second = datetime_match.groups()
                dt_obj = datetime(int(year), int(month), int(day), int(hour), int(minute), int(second))
                iso_formatted_date = dt_obj.strftime("%Y-%m-%dT%H:%M:%S.000")
                metadata['date'] = iso_formatted_date
                print(f"    Дата и время (из файла): {iso_formatted_date}")
            except ValueError: print("    Ошибка: Некорректная дата/время в имени файла.")
        else: print("    Предупреждение: Дата и время не найдены в имени файла.")


    metadata['client_name'] = "client"
    metadata['direction_outgoing'] = True
    metadata['language'] = "Ru-ru"

    if body:
        try:
            caller_match = re.search(r"Кто\sзвонил:\s*(.*?)(?=\nС\sкем\sговорил:|\nВремя\sзвонка:|\nДлительность:|\Z)", body, re.IGNORECASE | re.DOTALL)
            called_match = re.search(r"С\sкем\sговорил:\s*(.*?)(?=\nВремя\sзвонка:|\nДлительность:|\Z)", body, re.IGNORECASE | re.DOTALL)
            duration_match = re.search(
                r"Длительность:\s*(?:(\d+)\s*мин\.?)?(?:\s*(\d+)\s*сек)?", 
                body,
                re.IGNORECASE | re.DOTALL
            )

            if caller_match:
                caller_lines = caller_match.group(1).strip().split('\n')
                operator_name_raw = ' '.join(line.strip() for line in caller_lines if line.strip())
                metadata['operator_name'] = operator_name_raw
                metadata['operator_id'] = operator_name_raw.replace(' ', '_') if operator_name_raw else None
                print(f"    Оператор: {metadata.get('operator_name', 'Не найдено')}")
                print(f"    ID Оператора: {metadata.get('operator_id', 'Не найдено')}")
            else: print("    Предупреждение: Не удалось найти 'Кто звонил'"); metadata['operator_name'] = None; metadata['operator_id'] = None

            if called_match:
                 metadata['client_id'] = called_match.group(1).strip()
                 print(f"    ID Клиента (С кем говорил): {metadata.get('client_id', 'Не найдено')}")
            else: print("    Предупреждение: Не удалось найти 'С кем говорил'"); metadata['client_id'] = None


            # Логика парсинга длительности ---
            original_duration_str = "" 
            if duration_match:
                minutes_str = duration_match.group(1) # Группа 1: минуты (может быть None)
                seconds_str = duration_match.group(2) # Группа 2: секунды (может быть None)

                if minutes_str or seconds_str:
                    total_seconds = 0
                    parsed_successfully = True
                    try:
                        if minutes_str:
                            total_seconds += int(minutes_str) * 60
                            original_duration_str += f"{minutes_str} мин." 
                        if seconds_str:
                            total_seconds += int(seconds_str)
                            if original_duration_str: 
                                original_duration_str += " "
                            original_duration_str += f"{seconds_str} сек." 
                    except ValueError:
                        parsed_successfully = False
                        print(f"    Ошибка: Не удалось преобразовать минуты ('{minutes_str}') или секунды ('{seconds_str}') в число.")
                        original_duration_str_raw = duration_match.group(0).replace("Длительность:", "").strip()
                        metadata['call_duration_original_str'] = original_duration_str_raw
                    except Exception as e:
                         parsed_successfully = False
                         print(f"    Непредвиденная ошибка при обработке длительности: {e}")
                         original_duration_str_raw = duration_match.group(0).replace("Длительность:", "").strip()
                         metadata['call_duration_original_str'] = original_duration_str_raw

                    if parsed_successfully:
                        duration_seconds = total_seconds 
                        metadata['call_duration_parsed_sec'] = duration_seconds 
                        metadata['call_duration_original_str'] = original_duration_str.strip() 
                        print(f"    Длительность: {metadata['call_duration_original_str']} ({duration_seconds} секунд)")
                else:
                    print("    Предупреждение: Длительность найдена, но не удалось извлечь ни минуты, ни секунды.")
                    original_duration_str_raw = duration_match.group(0).replace("Длительность:", "").strip()
                    metadata['call_duration_original_str'] = original_duration_str_raw

            else:
                print("    Предупреждение: Не удалось найти 'Длительность' в ожидаемом формате.")

        except Exception as e:
            print(f"  Ошибка при парсинге метаданных из тела: {e}")
            if 'operator_name' not in metadata: metadata['operator_name'] = None
            if 'operator_id' not in metadata: metadata['operator_id'] = None
            if 'client_id' not in metadata: metadata['client_id'] = None
    else:
        print("  Тело письма пустое, метаданные из тела не извлечены.")
        metadata['operator_name'] = None; metadata['operator_id'] = None; metadata['client_id'] = None

    metadata = {k: v for k, v in metadata.items() if v is not None}
    return metadata, duration_seconds

if __name__ == "__main__":
    today = date.today(); imap_date_format = today.strftime("%d-%b-%Y"); folder_date_format = today.strftime("%d.%m.%Y")
    print(f"Дата поиска: {imap_date_format}, Папка даты: {folder_date_format}")
    SEARCH_CRITERIA = f'(ON "{imap_date_format}")'
    if not os.path.exists(DOWNLOAD_FOLDER): os.makedirs(DOWNLOAD_FOLDER); print(f"Создана папка: {DOWNLOAD_FOLDER}")
    date_folder_path = os.path.join(DOWNLOAD_FOLDER, folder_date_format)
    if not os.path.exists(date_folder_path): os.makedirs(date_folder_path); print(f"Создана папка: {date_folder_path}")

    mail = None
    processed_count = 0 
    try:
        print("Подключение..."); mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        print("Вход..."); mail.login(EMAIL_ACCOUNT, APP_PASSWORD); print("Успешный вход.")
        print(f"Выбор папки '{MAILBOX_FOLDER}'..."); status, messages_count_list = mail.select(MAILBOX_FOLDER)
        if status != 'OK': raise imaplib.IMAP4.error(f"Не удалось выбрать папку {MAILBOX_FOLDER}.")
        messages_count = messages_count_list[0].decode(); print(f"Папка '{MAILBOX_FOLDER}' выбрана, писем: {messages_count}")
        print(f"Поиск писем: {SEARCH_CRITERIA}..."); status, search_result = mail.search(None, SEARCH_CRITERIA)
        if status != 'OK': raise imaplib.IMAP4.error("Ошибка поиска.")
        email_ids = search_result[0].split()
        if not email_ids or email_ids == [b'']: print("Писем за сегодня нет."); email_ids = []
        else: print(f"Найдено писем: {len(email_ids)}")

        for email_id_bytes in email_ids:
            email_id_str = email_id_bytes.decode()
            print(f"\nОбработка письма ID: {email_id_str}")

            email_body = ""; found_mp3 = False; call_id = None
            call_folder_path_current = ""; original_mp3_filename = ""
            mp3_part_payload = None 

            try:
                status, msg_data = mail.fetch(email_id_bytes, '(RFC822)')
                if status != 'OK': print(f"  Не удалось получить письмо ID: {email_id_str}"); continue

                for response_part in msg_data:
                    if isinstance(response_part, tuple):
                        msg = email.message_from_bytes(response_part[1])

                        if not email_body:
                             if msg.is_multipart():
                                for part in msg.walk():
                                     ct = part.get_content_type(); cd = str(part.get("Content-Disposition"))
                                     if ct == "text/plain" and "attachment" not in cd:
                                          body_bytes = part.get_payload(decode=True)
                                          for cs in [part.get_content_charset(),'utf-8','cp1251','koi8-r','windows-1251']:
                                               if cs:
                                                   try: email_body = body_bytes.decode(cs, errors='ignore'); break
                                                   except: continue
                                          if not email_body: email_body = body_bytes.decode('utf-8', errors='ignore')
                                          if email_body: break 
                             else: 
                                ct = msg.get_content_type()
                                if ct == "text/plain":
                                     body_bytes = msg.get_payload(decode=True)
                                     for cs in [msg.get_content_charset(),'utf-8','cp1251','koi8-r','windows-1251']:
                                          if cs:
                                               try: email_body = body_bytes.decode(cs, errors='ignore'); break
                                               except: continue
                                     if not email_body: email_body = body_bytes.decode('utf-8', errors='ignore')

                        # --- Обработка вложений ---
                        if msg.is_multipart():
                             for part in msg.walk():
                                cd = str(part.get("Content-Disposition"))
                                if "attachment" in cd:
                                     filename_header = part.get_filename()
                                     original_filename = decode_filename(filename_header)

                                     if original_filename and original_filename.lower().endswith('.mp3'):
                                          found_mp3 = True
                                          original_mp3_filename = original_filename
                                          mp3_part_payload = part.get_payload(decode=True) 
                                          print(f"  Найдено вложение MP3: {original_mp3_filename}")



                if found_mp3:
                    call_id = extract_id_from_filename(original_mp3_filename)
                    if not call_id:
                         print("    Не удалось извлечь ID из имени файла, пропуск письма.")
                         continue 


                    metadata, duration_seconds = extract_metadata(email_body, call_id, folder_date_format, original_mp3_filename)

                    if duration_seconds > MIN_CALL_DURATION_SECONDS:
                        print(f"    Длительность ({duration_seconds} сек) > {MIN_CALL_DURATION_SECONDS} сек. Сохраняем звонок и метаданные.")


                        call_folder_name = call_id
                        call_folder_path_current = os.path.join(date_folder_path, call_folder_name)
                        if not os.path.exists(call_folder_path_current):
                            try:
                                os.makedirs(call_folder_path_current)
                                print(f"      Создана папка: {call_folder_path_current}")
                            except OSError as e:
                                print(f"      Ошибка создания папки ID {call_id}: {e}")
                                continue 


                        mp3_new_filename = f"{call_id}.mp3"
                        mp3_filepath = os.path.join(call_folder_path_current, mp3_new_filename)
                        try:
                            if not os.path.exists(mp3_filepath):
                                with open(mp3_filepath, 'wb') as f:
                                    f.write(mp3_part_payload) 
                                print(f"      Файл сохранен как '{mp3_new_filename}'.")
                            else:
                                print(f"      Файл '{mp3_new_filename}' уже существует.")
                        except Exception as e:
                            print(f"      Ошибка сохранения MP3 для ID {call_id}: {e}")
                            continue 

                        if metadata:
                            json_filepath = os.path.join(call_folder_path_current, METADATA_FILENAME)
                            try:
                                with open(json_filepath, 'w', encoding='utf-8') as f_json:
                                    json.dump(metadata, f_json, ensure_ascii=False, indent=4)
                                print(f"      Метаданные сохранены в '{json_filepath}'.")
                                processed_count += 1 # Увеличиваем счетчик
                            except Exception as e:
                                print(f"      Ошибка сохранения '{METADATA_FILENAME}' для ID {call_id}: {e}")
                        else:
                            print(f"      Метаданные для ID {call_id} не были извлечены.")

                    else:
                        print(f"    Длительность ({duration_seconds} сек) <= {MIN_CALL_DURATION_SECONDS} сек. Звонок пропущен.")
                else:
                    print("  Вложение MP3 не найдено в этом письме.")

            except Exception as e:
                 print(f"  Общая ошибка при обработке письма ID {email_id_str}: {e}")
                 continue

            # Опционально: Пометить письмо как прочитанное 
            # try: mail.store(email_id_bytes, '+FLAGS', '\\Seen') ...

    except imaplib.IMAP4.error as e: print(f"Ошибка IMAP: {e}")
    except Exception as e: print(f"Произошла непредвиденная ошибка: {e}")
    finally:
        # ... (код закрытия соединения) ...
        if mail and mail.state == 'SELECTED': mail.close()
        if mail: mail.logout(); print("\nСоединение закрыто.")

    print(f"\nРабота скрипта завершена. Обработано и сохранено звонков: {processed_count}")