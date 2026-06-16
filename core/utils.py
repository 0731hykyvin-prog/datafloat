import os
import re
from datetime import timedelta
import pandas as pd

def extract_phone_from_filename(file_path):
    """
    从文件名提取手机号

    示例:
    13055150066.xls
    13055150066[移动业务通话费详单]查询.xls
    移动_13055150066.xlsx
    """

    filename = os.path.basename(file_path)

    match = re.search(
        r'1[3-9]\d{9}',
        filename
    )

    if match:
        return match.group()

    return None
