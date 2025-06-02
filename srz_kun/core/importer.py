import csv
import os # For test file cleanup
from datetime import datetime

import openpyxl

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
    """Custom exception for errors related to import file format or content.

    This exception is raised when an imported CSV or Excel file does not conform
    to the expected structure, such as missing required headers, or when data
    values are invalid or cannot be normalized.
    """
    pass


def _normalize_person_data(row_dict: dict) -> dict:
    """Normalizes data from a row dictionary to match Person attribute requirements.

    Converts string values for boolean fields (like '生存'/'死亡' for `is_alive`
    and 'あり'/'なし' for `waived_inheritance`) to Python booleans.
    Validates and normalizes date string formats (YYYY-MM-DD, YYYY/MM/DD, YYYY.MM.DD)
    to "YYYY-MM-DD". Handles `datetime` objects if already converted by openpyxl.
    Clears `date_of_death` if `is_alive` is True.
    Ensures empty "ID" fields are set to `None`.

    Args:
        row_dict: A dictionary where keys are internal Person attribute names
            (e.g., 'name', 'is_alive', 'date_of_birth') and values are typically
            strings read from the source file, or `datetime` objects for dates
            from Excel.

    Returns:
        A dictionary with data normalized and typed appropriately for creating
        a `Person` object. For example, boolean strings are converted to
        Python booleans, and date strings are standardized.

    Raises:
        ImportFormatError: If essential data (name, relationship_to_deceased)
            is missing, or if boolean or date values have unrecognized formats.
    """
    normalized = row_dict.copy()
    internal_to_external_map = {v: k for k, v in EXPECTED_CSV_EXCEL_HEADERS.items()}

    if not normalized.get("name"):
        raise ImportFormatError(
            f"「{internal_to_external_map.get('name', '氏名')}」は必須項目です。"
        )
    if not normalized.get("relationship_to_deceased"):
        raise ImportFormatError(
            f"「{internal_to_external_map.get('relationship_to_deceased', '続柄')}」は必須項目です。"
        )

    is_alive_str = str(normalized.get("is_alive", "生存")).strip()
    if is_alive_str == "生存":
        normalized["is_alive"] = True
    elif is_alive_str == "死亡":
        normalized["is_alive"] = False
    else:
        raise ImportFormatError(
            f"「{internal_to_external_map.get('is_alive', '生死')}」の値が不正です: 「{is_alive_str}」。 "
            "「生存」または「死亡」を使用してください。"
        )

    waived_str = str(normalized.get("waived_inheritance", "なし")).strip()
    if waived_str == "あり":
        normalized["waived_inheritance"] = True
    elif waived_str == "なし":
        normalized["waived_inheritance"] = False
    else:
        raise ImportFormatError(
            f"「{internal_to_external_map.get('waived_inheritance', '相続放棄')}」の値が不正です: 「{waived_str}」。 "
            "「あり」または「なし」を使用してください。"
        )

    for date_field_internal_attr in ["date_of_birth", "date_of_death"]:
        date_val = normalized.get(date_field_internal_attr)
        user_facing_header = internal_to_external_map.get(
            date_field_internal_attr, date_field_internal_attr
        )

        if isinstance(date_val, str) and not date_val.strip():
            normalized[date_field_internal_attr] = None
        elif isinstance(date_val, str):
            date_val_stripped = date_val.strip()
            parsed_date = None
            # Try parsing common date formats
            for fmt_str, sep_char in [("%Y-%m-%d", "-"), ("%Y/%m/%d", "/"), ("%Y.%m.%d", ".")]:
                try:
                    parts = date_val_stripped.split(sep_char)
                    if len(parts) == 3 and all(p.isdigit() for p in parts): # Basic check
                        dt_obj = datetime.strptime(date_val_stripped, fmt_str)
                        parsed_date = dt_obj.strftime("%Y-%m-%d")
                        break  # Successfully parsed
                except ValueError:
                    continue # Try next format

            if parsed_date:
                normalized[date_field_internal_attr] = parsed_date
            else: # If no format matched
                raise ImportFormatError(
                    f"日付フィールド「{user_facing_header}」の値「{date_val}」は "
                    "YYYY-MM-DD (または YYYY/MM/DD, YYYY.MM.DD) 形式で入力してください。"
                )
        elif isinstance(date_val, datetime):  # From openpyxl
            normalized[date_field_internal_attr] = date_val.strftime("%Y-%m-%d")
        elif date_val is None:
            pass  # Keep as None
        else:
            raise ImportFormatError(
                f"日付フィールド「{user_facing_header}」の形式が不正です: 「{date_val}」。"
            )

    if normalized.get("is_alive") and normalized.get("date_of_death"):
        normalized["date_of_death"] = None

    if "id" in normalized and not str(normalized.get("id", "")).strip():
        normalized["id"] = None

    return normalized


def import_from_csv(filepath: str) -> list[dict]:
    """Imports person data from a CSV file.

    The CSV file must have a header row. Expected headers are defined in
    `EXPECTED_CSV_EXCEL_HEADERS`. "氏名" (name) and "続柄" (relationship)
    are required headers. Other fields are optional.

    Args:
        filepath: The path to the CSV file.

    Returns:
        A list of dictionaries, where each dictionary represents a person's
        data, normalized by `_normalize_person_data`.

    Raises:
        FileNotFoundError: If the CSV file does not exist at `filepath`.
        ImportFormatError: If the CSV format is invalid (e.g., missing headers,
            required data missing in a row, invalid values for boolean/date fields).
        Exception: For other unexpected errors during CSV processing.
    """
    people_data = []
    try:
        with open(filepath, mode='r', encoding='utf-8-sig', newline='') as csvfile:
            reader = csv.DictReader(csvfile)

            file_headers = reader.fieldnames
            if not file_headers:
                raise ImportFormatError("CSVファイルにヘッダー行が見つかりません。")

            for req_header in REQUIRED_EXTERNAL_HEADERS:
                if req_header not in file_headers:
                    raise ImportFormatError(
                        f"CSVファイルのヘッダーに必須項目「{req_header}」が見つかりません。"
                    )

            for i, row in enumerate(reader, 1):
                person_dict_raw = {}
                for external_header, internal_attr_name in EXPECTED_CSV_EXCEL_HEADERS.items():
                    # .get from DictReader row with default of empty string for missing optional columns
                    person_dict_raw[internal_attr_name] = row.get(external_header, "").strip() \
                        if row.get(external_header) is not None else ""

                # Skip rows that are likely empty or just commas
                if not any(person_dict_raw.get(attr) for attr in REQUIRED_EXTERNAL_HEADERS):
                    continue

                try:
                    normalized_data = _normalize_person_data(person_dict_raw)
                    people_data.append(normalized_data)
                except ImportFormatError as e:
                    raise ImportFormatError(f"CSVファイルの{i+1}行目（データ行{i}）でエラー: {e}")

    except FileNotFoundError:
        raise
    except csv.Error as e:
        raise ImportFormatError(f"CSVファイルの読み込み中にエラーが発生しました: {e}")
    except ImportFormatError:
        raise
    except Exception as e:
        raise Exception(f"CSVファイルの処理中に予期せぬエラーが発生しました: {e}")

    return people_data


def import_from_excel(filepath: str) -> list[dict]:
    """Imports person data from an Excel (.xlsx) file.

    Assumes data is in the first active sheet. The first row must be headers.
    Expected headers are defined in `EXPECTED_CSV_EXCEL_HEADERS`.
    "氏名" (name) and "続柄" (relationship) are required headers.

    Args:
        filepath: The path to the Excel file.

    Returns:
        A list of dictionaries, where each dictionary represents a person's
        data, normalized by `_normalize_person_data`.

    Raises:
        FileNotFoundError: If the Excel file does not exist at `filepath`.
        ImportFormatError: If the Excel format is invalid (e.g., .xlsx format error,
            missing headers, required data missing, invalid values).
        Exception: For other unexpected errors during Excel processing.
    """
    people_data = []
    try:
        workbook = openpyxl.load_workbook(filepath, data_only=True)
        sheet = workbook.active

        if sheet.max_row < 1:
            raise ImportFormatError("Excelファイルが空か、ヘッダー行がありません。")

        header_row_iter = sheet.iter_rows(min_row=1, max_row=1, values_only=True)
        header_row_values = next(header_row_iter, None)
        if not header_row_values:
             raise ImportFormatError("Excelファイルからヘッダー行を読み取れませんでした。")

        header_row = [str(h).strip() if h is not None else "" for h in header_row_values]

        for req_header in REQUIRED_EXTERNAL_HEADERS:
            if req_header not in header_row:
                raise ImportFormatError(
                    f"Excelファイルのヘッダーに必須項目「{req_header}」が見つかりません。"
                )

        for i, row_values_tuple in enumerate(sheet.iter_rows(min_row=2, values_only=True), 2):
            if not any(row_values_tuple):
                continue

            person_dict_raw = {}
            for external_header, internal_attr_name in EXPECTED_CSV_EXCEL_HEADERS.items():
                try:
                    col_idx = header_row.index(external_header)
                    cell_val = row_values_tuple[col_idx]

                    if isinstance(cell_val, datetime):
                        person_dict_raw[internal_attr_name] = cell_val # Pass datetime directly to normalize
                    elif cell_val is not None:
                        person_dict_raw[internal_attr_name] = str(cell_val).strip()
                    else:
                        person_dict_raw[internal_attr_name] = None
                except (ValueError, IndexError):
                    person_dict_raw[internal_attr_name] = None

            try:
                normalized_data = _normalize_person_data(person_dict_raw)
                people_data.append(normalized_data)
            except ImportFormatError as e:
                raise ImportFormatError(f"Excelファイルの{i}行目でエラー: {e}")

    except FileNotFoundError:
        raise
    except openpyxl.utils.exceptions.InvalidFileException:
        raise ImportFormatError(
            f"Excelファイル形式が無効です。xlsx形式のファイルのみ対応しています: {filepath}"
        )
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
        # Using datetime objects for dates in Excel as openpyxl handles them well
        ws.append(["", "山田 太郎", "被相続人", datetime(1950,1,1), datetime(2023,10,26), "東京都千代田区", "東京都新宿区", "死亡", "なし"])
        ws.append(["id002", "山田 花子", "妻", datetime(1955,5,10), None, "東京都千代田区", "東京都新宿区", "生存", "なし"])
        ws.append(["", "山田 一郎", "長男", datetime(1980,8,15), None, "東京都新宿区", "東京都新宿区", "生存", "あり"])
        # For testing string date parsing from Excel, one could also write "1990/03/20" as a string
        ws.append(["", "佐藤 スミス", "養子", "1990-03-20", None, "神奈川県横浜市", "神奈川県横浜市", "生存", "なし"])
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
            writer.writerow(["テスト 太郎", "本人", "01-01-2000"]) # Invalid format not in normalization list
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
