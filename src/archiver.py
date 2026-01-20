import os
import shutil

def archive(src, base_dir, category, new_name):
    # 目标目录
    dst_dir = os.path.join(base_dir, category)
    os.makedirs(dst_dir, exist_ok=True)

    # 文件扩展名
    ext = os.path.splitext(src)[1]

    # 最终路径
    dst = os.path.join(dst_dir, new_name + ext)

    # 再次确保路径合法
    dst = dst.replace(':', '_').replace('*', '_').replace('?', '_')

    # 复制
    shutil.copy2(src, dst)
    return dst
