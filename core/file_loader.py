import os

SUPPORTED_EXTENSIONS = {".xlsx", ".xls", ".xlsm", ".csv"}


def scan_files(folder):
    result = []

    for root, dirs, files in os.walk(folder):
        for file in files:
            if file.startswith("~$"):
                continue

            ext = os.path.splitext(file)[1].lower()
            if ext in SUPPORTED_EXTENSIONS:
                result.append(os.path.join(root, file))

    return result
