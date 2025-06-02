import csv
import openpyxl
from datetime import datetime # For date validation/conversion if needed
import os # For test file cleanup

# Define expected headers and their mapping to Person object attribute names
# These are the headers we expect in the CSV/Excel files.
EXPECTED_CSV_EXCEL_HEADERS = {
    "ID": "id",
    "氏名": "name",
    "続柄": "relationship_to_deceased",
    "生年月日": "date_of_birth",
    "死亡年月日": "date_of_death",
    "本籍地": "permanent_domicile",
    "住所": "address",
    "生死": "is_alive",  # Expected values: "生存", "死亡"
    "相続放棄": "waived_inheritance"  # Expected values: "あり", "なし"
}
# Required external headers for a row to be considered valid
REQUIRED_EXTERNAL_HEADERS = ["氏名", "続柄"]


class ImportFormatError(ValueError):
    """Custom exception for import format errors."""
    pass


def _normalize_person_data(row_dict):
    """
    Normalizes data from a row dictionary (from CSV/Excel) to match Person attributes.
    Converts string values like '生存'/'死亡' and 'あり'/'なし' to booleans.
    Validates date formats if possible.
    Args:
        row_dict (dict): A dictionary where keys are internal Person attribute names
                         (e.g., 'name', 'is_alive') and values are strings from the file.
    Returns:
        dict: A dictionary with normalized data.
    Raises:
        ImportFormatError: If essential data is missing or in an unexpected format.
    """
    normalized = row_dict.copy()

    # Use a reverse map for error messages to show user-facing header names
    internal_to_external_map = {v: k for k, v in EXPECTED_CSV_EXCEL_HEADERS.items()}


    # Name and Relationship are critical and should be present from header mapping
    if not normalized.get("name"):
        raise ImportFormatError(f"「{internal_to_external_map.get('name', '氏名')}」は必須項目です。")
    if not normalized.get("relationship_to_deceased"):
        raise ImportFormatError(f"「{internal_to_external_map.get('relationship_to_deceased', '続柄')}」は必須項目です。")

    # Normalize boolean fields
    is_alive_str = str(normalized.get("is_alive", "生存")).strip() # Default to alive, ensure string
    if is_alive_str == "生存":
        normalized["is_alive"] = True
    elif is_alive_str == "死亡":
        normalized["is_alive"] = False
        # No specific check for date_of_death here, but it will be processed next
    else:
        raise ImportFormatError(f"「{internal_to_external_map.get('is_alive', '生死')}」の値が不正です: 「{is_alive_str}」。 「生存」または「死亡」を使用してください。")

    waived_str = str(normalized.get("waived_inheritance", "なし")).strip() # Default to 'なし', ensure string
    if waived_str == "あり":
        normalized["waived_inheritance"] = True
    elif waived_str == "なし":
        normalized["waived_inheritance"] = False
    else:
        raise ImportFormatError(f"「{internal_to_external_map.get('waived_inheritance', '相続放棄')}」の値が不正です: 「{waived_str}」。 「あり」または「なし」を使用してください。")

    # Normalize date fields
    for date_field_internal_attr in ["date_of_birth", "date_of_death"]:
        date_val = normalized.get(date_field_internal_attr)
        user_facing_header = internal_to_external_map.get(date_field_internal_attr, date_field_internal_attr)

        if isinstance(date_val, str) and not date_val.strip():
            normalized[date_field_internal_attr] = None
        elif isinstance(date_val, str):
            date_val_stripped = date_val.strip()
            parsed_date = None
            for fmt_str, sep in [("%Y-%m-%d", "-"), (f"%Y/%m/%d", "/"), (f"%Y.%m.%d", ".")]:
                try:
                    # Check if all parts are digits before trying to parse
                    parts = date_val_stripped.split(sep)
                    if len(parts) == 3 and all(p.isdigit() for p in parts):
                        dt_obj = datetime.strptime(date_val_stripped, fmt_str)
                        parsed_date = dt_obj.strftime("%Y-%m-%d")
                        break
                except ValueError:
                    continue
            if parsed_date:
                 normalized[date_field_internal_attr] = parsed_date
            else:
                raise ImportFormatError(
                    f"日付フィールド「{user_facing_header}」の値「{date_val}」は "
                    f"YYYY-MM-DD (または YYYY/MM/DD, YYYY.MM.DD) 形式で入力してください。"
                )
        elif isinstance(date_val, datetime): # Already a datetime object (e.g., from openpyxl)
            normalized[date_field_internal_attr] = date_val.strftime("%Y-%m-%d")
        elif date_val is None:
            pass # Already None
        else: # Some other unexpected type
            raise ImportFormatError(
                f"日付フィールド「{user_facing_header}」の形式が不正です: 「{date_val}」。"
            )

    # Ensure date_of_death is None if person is alive
    if normalized.get("is_alive") and normalized.get("date_of_death"):
        # print(f"Info: Person '{normalized.get('name')}' is alive but had a date of death; DoD cleared.")
        normalized["date_of_death"] = None

    # ID can be None or empty string, will be handled by Person class
    if "id" in normalized and not str(normalized["id"]).strip(): # Ensure ID is string before strip
        normalized["id"] = None

    return normalized


def import_from_csv(filepath):
    people_data = []
    try:
        with open(filepath, mode='r', encoding='utf-8-sig', newline='') as csvfile:
            reader = csv.DictReader(csvfile)

            file_headers = reader.fieldnames
            if not file_headers:
                raise ImportFormatError("CSVファイルにヘッダー行が見つかりません。")

            for req_header in REQUIRED_EXTERNAL_HEADERS:
                if req_header not in file_headers:
                    raise ImportFormatError(f"CSVファイルのヘッダーに必須項目「{req_header}」が見つかりません。")

            for i, row in enumerate(reader, 1): # Start row count from 1 (after header)
                person_dict_raw = {}
                for external_header, internal_attr_name in EXPECTED_CSV_EXCEL_HEADERS.items():
                    person_dict_raw[internal_attr_name] = row.get(external_header, "").strip() if row.get(external_header) is not None else ""

                try:
                    # Skip entirely empty rows more reliably
                    if not any(person_dict_raw.get(attr) for attr in ["name", "relationship_to_deceased", "date_of_birth", "date_of_death"]): # check some key fields
                        continue
                    normalized_data = _normalize_person_data(person_dict_raw)
                    people_data.append(normalized_data)
                except ImportFormatError as e:
                    raise ImportFormatError(f"CSVファイルの{i+1}行目（データ行{i}）でエラー: {e}")

    except FileNotFoundError:
        raise FileNotFoundError(f"ファイルが見つかりません: {filepath}")
    except csv.Error as e: # More generic CSV error
        raise ImportFormatError(f"CSVファイルの読み込み中にエラーが発生しました: {e}")
    except ImportFormatError:
        raise
    except Exception as e:
        raise Exception(f"CSVファイルの処理中に予期せぬエラーが発生しました: {e}")

    return people_data


def import_from_excel(filepath):
    people_data = []
    try:
        workbook = openpyxl.load_workbook(filepath, data_only=True)
        sheet = workbook.active

        if sheet.max_row < 1:
            raise ImportFormatError("Excelファイルが空か、ヘッダー行がありません。")

        header_row_iter = sheet.iter_rows(min_row=1, max_row=1, values_only=True)
        header_row = next(header_row_iter, None)
        if not header_row:
             raise ImportFormatError("Excelファイルからヘッダー行を読み取れませんでした。")

        header_row = [str(h).strip() if h is not None else "" for h in header_row]


        for req_header in REQUIRED_EXTERNAL_HEADERS:
            if req_header not in header_row:
                raise ImportFormatError(f"Excelファイルのヘッダーに必須項目「{req_header}」が見つかりません。")

        for i, row_values_tuple in enumerate(sheet.iter_rows(min_row=2, values_only=True), 2): # Start from 2nd row
            if not any(row_values_tuple): # Skip entirely empty rows
                continue

            person_dict_raw = {}
            for external_header, internal_attr_name in EXPECTED_CSV_EXCEL_HEADERS.items():
                try:
                    col_idx = header_row.index(external_header)
                    cell_val = row_values_tuple[col_idx]
                    # openpyxl might return datetime objects for dates
                    if isinstance(cell_val, datetime):
                        person_dict_raw[internal_attr_name] = cell_val.strftime("%Y-%m-%d")
                    elif cell_val is not None:
                        person_dict_raw[internal_attr_name] = str(cell_val).strip()
                    else:
                        person_dict_raw[internal_attr_name] = None # Explicitly set to None
                except (ValueError, IndexError):
                    person_dict_raw[internal_attr_name] = None

            try:
                normalized_data = _normalize_person_data(person_dict_raw)
                people_data.append(normalized_data)
            except ImportFormatError as e:
                raise ImportFormatError(f"Excelファイルの{i}行目でエラー: {e}")

    except FileNotFoundError:
        raise FileNotFoundError(f"ファイルが見つかりません: {filepath}")
    except openpyxl.utils.exceptions.InvalidFileException:
        raise ImportFormatError(f"Excelファイル形式が無効です。xlsx形式のファイルのみ対応しています: {filepath}")
    except ImportFormatError:
        raise
    except Exception as e:
        raise Exception(f"Excelファイルの処理中に予期せぬエラーが発生しました: {e}")

    return people_data

if __name__ == '__main__':
    # Basic test examples

    dummy_csv_path = "dummy_test_people.csv"
    dummy_excel_path = "dummy_test_people.xlsx"
    bad_date_csv_path = "bad_date.csv"
    missing_header_csv_path = "missing_header.csv"

    try:
        print("\n--- CSV Import Test ---")
        with open(dummy_csv_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(["ID", "氏名", "続柄", "生年月日", "死亡年月日", "本籍地", "住所", "生死", "相続放棄"])
            writer.writerow(["", "山田 太郎", "被相続人", "1950-01-01", "2023-10-26", "東京都千代田区", "東京都新宿区", "死亡", "なし"])
            writer.writerow(["id002", "山田 花子", "妻", "1955/05/10", "", "東京都千代田区", "東京都新宿区", "生存", "なし"])
            writer.writerow(["", "山田 一郎", "長男", "1980.08.15", "", "東京都新宿区", "東京都新宿区", "生存", "あり"])
            writer.writerow(["", "佐藤 スミス", "養子", "1990-03-20", "", "神奈川県横浜市", "神奈川県横浜市", "生存", "なし"])

        csv_data = import_from_csv(dummy_csv_path)
        print(f"CSV Imported {len(csv_data)} people.")
        for person in csv_data:
            print(person)
    except Exception as e:
        print(f"CSV Test Error: {e}")
    finally:
        if os.path.exists(dummy_csv_path): os.remove(dummy_csv_path)

    try:
        print("\n--- Excel Import Test ---")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["ID", "氏名", "続柄", "生年月日", "死亡年月日", "本籍地", "住所", "生死", "相続放棄"])
        ws.append(["", "山田 太郎", "被相続人", datetime(1950,1,1), datetime(2023,10,26), "東京都千代田区", "東京都新宿区", "死亡", "なし"])
        ws.append(["id002", "山田 花子", "妻", "1955-05-10", None, "東京都千代田区", "東京都新宿区", "生存", "なし"])
        ws.append(["", "山田 一郎", "長男", "1980/08/15", None, "東京都新宿区", "東京都新宿区", "生存", "あり"])
        ws.append(["", "佐藤 スミス", "養子", "1990.03.20", None, "神奈川県横浜市", "神奈川県横浜市", "生存", "なし"])
        wb.save(dummy_excel_path)

        excel_data = import_from_excel(dummy_excel_path)
        print(f"Excel Imported {len(excel_data)} people.")
        for person in excel_data:
            print(person)
    except Exception as e:
        print(f"Excel Test Error: {e}")
    finally:
        if os.path.exists(dummy_excel_path): os.remove(dummy_excel_path)

    try:
        print("\n--- Bad Date Format Test (CSV) ---")
        with open(bad_date_csv_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(["氏名", "続柄", "生年月日"])
            writer.writerow(["テスト 太郎", "本人", "01-01-2000"])
        import_from_csv(bad_date_csv_path)
    except ImportFormatError as e:
        print(f"Bad Date Test Caught Expected Error: {e}")
    except Exception as e:
        print(f"Bad Date Test Caught Unexpected Error: {e}")
    finally:
        if os.path.exists(bad_date_csv_path): os.remove(bad_date_csv_path)

    try:
        print("\n--- Missing Header Test (CSV) ---")
        with open(missing_header_csv_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(["氏名", "生年月日"])
            writer.writerow(["テスト 花子", "2000-01-01"])
        import_from_csv(missing_header_csv_path)
    except ImportFormatError as e:
        print(f"Missing Header Test Caught Expected Error: {e}")
    except Exception as e:
        print(f"Missing Header Test Caught Unexpected Error: {e}")
    finally:
        if os.path.exists(missing_header_csv_path): os.remove(missing_header_csv_path)
```
