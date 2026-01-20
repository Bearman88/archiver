import os
import yaml
from src.extractor import extract_text
from src.classifier import classify
from src.renamer import build_filename
from src.archiver import archive
from src.logger import Logger

from pdf2image import convert_from_path
from pdf2image.exceptions import PDFPageCountError
import pytesseract
import comtypes.client

# ==========================
# 配置
# ==========================
SOURCE_DIR = r'D:\BaiduSyncdisk\1【品种】天气\2023\新闻'
ARCHIVE_DIR = r'D:\BaiduSyncdisk\1【品种】天气\2023\新闻\归档'

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
POPPLER_PATH = r"C:\poppler-25.12.0\Library\bin"

with open('config/rules.yaml', 'r', encoding='utf-8') as f:
    rules = yaml.safe_load(f)

logger = Logger()

# --------------------------
# DOC/WPS 转 DOCX
# --------------------------
def convert_to_docx(src_path):
    ext = os.path.splitext(src_path)[1].lower()
    if ext not in ['.doc', '.wps']:
        return src_path
    word = comtypes.client.CreateObject('Word.Application')
    word.Visible = False
    doc = word.Documents.Open(src_path)
    dst_path = src_path + 'x' if ext == '.doc' else src_path.replace('.wps', '.docx')
    doc.SaveAs(dst_path, FileFormat=16)
    doc.Close()
    word.Quit()
    return dst_path

# --------------------------
# OCR 提取 PDF / 图片标题（安全版）
# --------------------------
def ocr_extract_first_line(path):
    ext = os.path.splitext(path)[1].lower()
    text = ""
    try:
        if ext == '.pdf':
            pages = convert_from_path(path, dpi=200, poppler_path=POPPLER_PATH)
            if pages:
                text = pytesseract.image_to_string(pages[0], lang='chi_sim')
        elif ext in ['.png', '.jpg', '.jpeg']:
            text = pytesseract.image_to_string(path, lang='chi_sim')
    except PDFPageCountError:
        logger.fail(文件=path, 原因='PDF 损坏或非法', 内容预览='')
        return None
    except Exception as e:
        logger.fail(文件=path, 原因=f'OCR 错误: {e}', 内容预览='')
        return None

    for line in text.splitlines():
        if line.strip():
            return line.strip()
    return os.path.splitext(os.path.basename(path))[0]

# --------------------------
# 扫描所有文件
# --------------------------
for root, _, files in os.walk(SOURCE_DIR):
    for file in files:
        path = os.path.join(root, file)
        ext = os.path.splitext(file)[1].lower()

        # DOC/WPS 先转换
        if ext in ['.doc', '.wps']:
            try:
                path = convert_to_docx(path)
                ext = os.path.splitext(path)[1].lower()
            except Exception as e:
                logger.fail(文件=path, 原因=f'DOC/WPS 转 DOCX 失败: {e}', 内容预览=file)
                continue  # 转换失败跳过

        # 特殊分组
        handled_special = False
        for grp, cfg in rules.get('special_groups', {}).items():
            if ext in cfg['extensions']:
                title = os.path.splitext(file)[0]
                new_name = build_filename(path, title)
                try:
                    dst = archive(path, ARCHIVE_DIR, grp, new_name)
                    logger.log(
                        原路径=path,
                        新路径=dst,
                        分类=grp,
                        文件类型=ext,
                        重命名结果=new_name,
                        置信分=999
                    )
                except Exception as e:
                    logger.fail(文件=path, 原因=str(e), 内容预览=title)
                handled_special = True
                break
        if handled_special:
            continue

        # OCR / 标题提取
        if ext in ['.pdf', '.png', '.jpg', '.jpeg']:
            title = ocr_extract_first_line(path)
            if title is None:  # OCR 失败直接跳过
                continue
        else:
            lines = extract_text(path)
            title = lines[0] if lines else "未命名"

        # 分类
        try:
            if ext not in ['.pdf', '.png', '.jpg', '.jpeg']:
                lines = extract_text(path)
                cat, score = classify(lines, rules)
            else:
                cat, score = classify([title], rules)
        except Exception:
            cat, score = rules['fallback']['name'], 0

        if not cat:
            cat = rules['fallback']['name']
            logger.fail(文件=path, 原因='未命中规则', 内容预览=title)

        # Windows 安全命名 + 归档
        new_name = build_filename(path, title)
        new_name = new_name.replace(':', '_').replace('*', '_').replace('?', '_')\
                           .replace('"', '_').replace('<', '_').replace('>', '_')\
                           .replace('|', '_').replace('\\','_').replace('/','_')

        try:
            dst = archive(path, ARCHIVE_DIR, cat, new_name)
            logger.log(
                原路径=path,
                新路径=dst,
                分类=cat,
                置信分=score,
                文件类型=ext,
                重命名结果=new_name
            )
        except Exception as e:
            logger.fail(文件=path, 原因=str(e), 内容预览=title)

logger.save()
print("归档完成，清单和失败日志已生成在 output/ 下。")
