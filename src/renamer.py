import os
import re
from datetime import datetime

def build_filename(path, title):
    # 获取最后修改日期
    try:
        mtime = os.path.getmtime(path)
        date_str = datetime.fromtimestamp(mtime).strftime("%Y%m%d")
    except:
        date_str = datetime.now().strftime("%Y%m%d")

    # Windows 文件名安全化
    if not title or not title.strip():
        safe_title = "未命名"
    else:
        safe_title = re.sub(r'[\\/:*?"<>|]', '_', title)
        safe_title = safe_title.strip()[:100]  # 限制长度避免过长

    return f"{date_str}_{safe_title}"
