import os
import shutil
import zipfile
import glob
import openpyxl
from datetime import datetime, timedelta
import time
import sys
import re
import pandas as pd
from mail_notifier import send_mail

# ==================== 配置文件读取 ====================
def load_config():
    """
    从与程序同目录的 config.txt 读取配置（键=值格式）。
    若文件不存在，返回默认值。
    """
    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.txt")
    defaults = {
        "MONITOR_DIR": r"D:\exercise\python\pthto",
        "ZIP_PATTERN": "*xg简易自动巡检系统*照片-昨日*.zip",
        "FIXED_ZIP": "",                     # 留空则使用监控目录，非空则优先使用
        "WORK_DIR": r"D:\exercise\python\temp",
        "OUTPUT_DIR": r"D:\exercise\python\output",
        "SCHEDULE_TIME": "",                 # 留空则立即执行，可填 "02:00"
        "LOG_DIR": "D:\exercise\python\logs",                       # 留空则使用 OUTPUT_DIR
        "CLEAN_TEMP": "1",  # 新增：1 表示自动清理临时目录，0 表示不清理
        # "FINAL_DIR": "",
        "FINAL_DIR": r"E:\光伏运维\01流水资料\自动下载目录\照片重命名分类后",   # 最终归档目录，留空则不搬运。在字符串前加 r，禁止转义
        #以下是邮件部分
        "ENABLE_MAIL": "0",
        "SMTP_SERVER": "smtp.139.com",
        "SMTP_PORT": "465",
        "SMTP_USER": "",
        "SMTP_PASSWORD": "",
        "MAIL_TO": "",
        "MAIL_SUBJECT_SUCCESS": "巡检照片处理成功",
        "MAIL_SUBJECT_FAIL": "巡检照片处理失败",
    }
    if not os.path.exists(config_path):
        return defaults

    config = defaults.copy()
    with open(config_path, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith('#'):
                continue
            if '=' in line:
                key, value = line.split('=', 1)
                key = key.strip()
                value = value.strip().strip('"').strip("'")
                if key in config:
                    config[key] = value
    return config

# 加载配置
cfg = load_config()
MONITOR_DIR = cfg.get("MONITOR_DIR", "")
ZIP_PATTERN = cfg.get("ZIP_PATTERN", "*xg简易自动巡检系统*照片-昨日*.zip")
FIXED_ZIP = cfg.get("FIXED_ZIP", "") or None
WORK_DIR = cfg.get("WORK_DIR", r"D:\exercise\python\temp")
OUTPUT_DIR = cfg.get("OUTPUT_DIR", r"D:\exercise\python\output")
SCHEDULE_TIME = cfg.get("SCHEDULE_TIME", "") or None
LOG_DIR = cfg.get("LOG_DIR", "") or OUTPUT_DIR
CLEAN_TEMP = cfg.get("CLEAN_TEMP", "0") == "1"   # 字符串"1"才视为清理
FINAL_DIR = cfg.get("FINAL_DIR", "") or None

ENABLE_MAIL = cfg.get("ENABLE_MAIL", "0") == "1"
SMTP_SERVER = cfg.get("SMTP_SERVER", "")
SMTP_PORT = cfg.get("SMTP_PORT", "587")
SMTP_USER = cfg.get("SMTP_USER", "")
SMTP_PASSWORD = cfg.get("SMTP_PASSWORD", "")
MAIL_TO = cfg.get("MAIL_TO", "")
MAIL_SUBJECT_SUCCESS = cfg.get("MAIL_SUBJECT_SUCCESS", "巡检照片处理成功")
MAIL_SUBJECT_FAIL = cfg.get("MAIL_SUBJECT_FAIL", "巡检照片处理失败")

# 光伏区关键词
SOLAR_RULES = [
    "光伏区-进出口标识", "进出入口标识", "光伏区-进出入口通道",
    "光伏区-屋面安全防护", "光伏区-屋面天沟", "光伏区-屋面污染情况",
    "光伏区-组件", "逆变器", "汇流箱", "电缆桥架", "光伏区-桥架",
    "线缆", "光伏区-线缆", "水管", "消防", "光伏区-设备及通道",
    "光伏区-巡视完成后",
]

# 配电区关键词
DISTRIBUTION_RULES = [
    "低压配电室", "高压配电室", "低配区", "高配区",
]

PHOTO_EXT = ('.jpg', '.jpeg', '.png', '.bmp')


# ==================== 工具函数 ====================
def log(msg):
    """打印并写入日志文件（按日期命名）"""
    now = datetime.now()
    time_str = now.strftime('%H:%M:%S')
    print(f"[{time_str}] {msg}")

    # 写入文件
    log_filename = os.path.join(LOG_DIR, f"process_{now.strftime('%Y%m%d')}.log")
    os.makedirs(LOG_DIR, exist_ok=True)
    with open(log_filename, 'a', encoding='utf-8') as f:
        f.write(f"[{now.strftime('%Y-%m-%d %H:%M:%S')}] {msg}\n")


def wait_until_scheduled_time(schedule_time_str):
    """
    等待直到当天的指定时间（格式 HH:MM）。
    若当前时间已过该时刻，则等待到第二天的同一时刻。
    """
    now = datetime.now()
    target_time = datetime.strptime(schedule_time_str, "%H:%M").time()
    target_datetime = datetime.combine(now.date(), target_time)
    if target_datetime <= now:
        target_datetime += timedelta(days=1)   # 等到明天
    wait_seconds = (target_datetime - now).total_seconds()
    log(f"当前时间 {now.strftime('%H:%M:%S')}，将等待至 {target_datetime.strftime('%Y-%m-%d %H:%M:%S')} 开始执行（约 {wait_seconds/60:.1f} 分钟）")
    time.sleep(wait_seconds)

def check_directories(config):
    """检查配置中的关键目录是否存在，不存在则发出警告"""
    for key in ['MONITOR_DIR', 'WORK_DIR', 'OUTPUT_DIR', 'LOG_DIR']:
        path = config.get(key, '')
        if path and not os.path.isdir(path):
            log(f"⚠️ 配置路径不存在: {key} = {path}（程序可能会因此失败）")

def find_latest_zip():
    if not MONITOR_DIR or not os.path.isdir(MONITOR_DIR):
        log("未配置有效的监控目录，尝试使用 FIXED_ZIP")
        return FIXED_ZIP

    search_pattern = os.path.join(MONITOR_DIR, ZIP_PATTERN)
    files = glob.glob(search_pattern)
    if not files:
        log(f"在 {MONITOR_DIR} 下未找到匹配 {ZIP_PATTERN} 的文件，尝试使用 FIXED_ZIP")
        return FIXED_ZIP

    files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    latest = files[0]
    log(f"自动选择最新压缩包：{latest}")
    return latest

def find_all_zips():
    """在监控目录下按通配模式查找所有匹配的 ZIP 文件，返回按修改时间排序的路径列表"""
    if not MONITOR_DIR or not os.path.isdir(MONITOR_DIR):
        log("未配置有效的监控目录，尝试使用 FIXED_ZIP")
        if FIXED_ZIP:
            return [FIXED_ZIP]
        else:
            return []

    search_pattern = os.path.join(MONITOR_DIR, ZIP_PATTERN)
    files = glob.glob(search_pattern)
    if not files:
        log(f"在 {MONITOR_DIR} 下未找到匹配 {ZIP_PATTERN} 的文件")
        if FIXED_ZIP:
            return [FIXED_ZIP]
        else:
            return []

    files.sort(key=lambda x: os.path.getmtime(x))
    log(f"共找到 {len(files)} 个压缩包待处理：{files}")
    return files

# ==================== 1. 复制并解压 ====================
def unzip_file(zip_path):
    if not os.path.exists(zip_path):
        raise FileNotFoundError(f"找不到压缩包：{zip_path}")
    os.makedirs(WORK_DIR, exist_ok=True)
    zip_name = os.path.basename(zip_path)
    dest_zip = os.path.join(WORK_DIR, zip_name)
    shutil.copy2(zip_path, dest_zip)
    log(f"已复制压缩包到：{dest_zip}")
    extract_root = os.path.join(WORK_DIR, "extracted")
    os.makedirs(extract_root, exist_ok=True)
    with zipfile.ZipFile(dest_zip, 'r') as z:
        z.extractall(extract_root)
    log(f"解压完成，路径：{extract_root}")
    return extract_root


# ==================== 2. 获取站点名称与数据时间 ====================
def get_site_and_date(zip_path):
    basename = os.path.splitext(os.path.basename(zip_path))[0]
    # 取最后一个下划线之前的部分作为站点前缀
    site = basename.rsplit('_', 1)[0] if '_' in basename else basename
    # 从文件修改时间获取日期
    mtime = os.path.getmtime(zip_path)
    date_str = datetime.fromtimestamp(mtime).strftime("%Y%m%d")
    log(f"识别站点前缀：{site}，日期（文件时间）：{date_str}")
    return site, date_str


# ==================== 3. 查找配套 Excel 文件 ====================
def find_excel_file(zip_path):
    """根据压缩包路径，在同一目录下查找前缀相同的 xlsx 文件"""
    zip_dir = os.path.dirname(zip_path)
    zip_basename = os.path.basename(zip_path)
    # 去掉 .zip 扩展名
    zip_name_noext = os.path.splitext(zip_basename)[0]
    # 提取公共前缀：到最后一个下划线之前（去掉 _附件(数字) 部分）
    # 例如 "DH光伏巡检-照片巡检_xg简易自动巡检系统_照片-昨日_附件(2)" -> 前缀为 "DH光伏巡检-照片巡检_xg简易自动巡检系统_照片-昨日"
    prefix = zip_name_noext.rsplit('_', 1)[0] if '_' in zip_name_noext else zip_name_noext

    # 使用 glob 模糊匹配
    search_pattern = os.path.join(zip_dir, prefix + "*.xlsx")
    matches = glob.glob(search_pattern)
    if matches:
        # 如果有多个，取修改时间最新的（或任意）
        matches.sort(key=lambda x: os.path.getmtime(x), reverse=True)
        excel_path = matches[0]
        log(f"找到配套 Excel：{excel_path}")
        return excel_path
    else:
        log(f"⚠️ 未找到匹配的 Excel 文件，模式：{search_pattern}")
        return None

# ==================== 4. 从 Excel 构建文件名 -> 站点映射 ====================
def build_filename_site_map(excel_path):
    """
    使用 pandas 解析 Excel（容错性更强），返回 {照片完整文件名（小写）: [站点名称列表]}
    """
    # 读取整个 Excel 为 DataFrame，不设表头，稍后手动提取
    df = pd.read_excel(excel_path, header=None, engine='openpyxl')

    # 第一行就是表头
    headers = df.iloc[0].tolist()
    log(f"Excel 表头内容：{headers}")

    # 查找“站点名称”列
    site_col = None
    possible_names = ['站点名称', '站点', '项目名称', '电站名称']
    for idx, h in enumerate(headers):
        if h and any(name in str(h) for name in possible_names):
            site_col = idx
            log(f"找到站点列：“{h}” (索引 {idx})")
            break
    if site_col is None:
        raise ValueError(f"未找到站点名称列，表头为：{headers}")

    # 查找所有包含“巡检照片”或“照片”的列
    photo_cols = [idx for idx, h in enumerate(headers) if h and ('巡检照片' in str(h) or '照片' in str(h))]
    if not photo_cols:
        raise ValueError(f"未找到巡检照片列，表头为：{headers}")

    mapping = {}
    # 从第2行开始遍历数据
    for i in range(1, len(df)):
        row = df.iloc[i]
        site_name = row.iloc[site_col] if pd.notna(row.iloc[site_col]) else None
        if not site_name:
            continue
        for col_idx in photo_cols:
            cell_val = row.iloc[col_idx] if pd.notna(row.iloc[col_idx]) else None
            if cell_val:
                filenames = [f.strip() for f in str(cell_val).split(',') if f.strip()]
                for fn in filenames:
                    key = fn.lower()
                    if key not in mapping:
                        mapping[key] = []
                    if site_name not in mapping[key]:
                        mapping[key].append(site_name)

    # 统计歧义
    multi_site_count = sum(1 for v in mapping.values() if len(v) > 1)
    log(f"从 Excel 解析出 {len(mapping)} 个照片文件名，其中 {multi_site_count} 个存在多站点歧义")
    if multi_site_count > 0:
        log("歧义文件名示例：")
        for k, v in mapping.items():
            if len(v) > 1:
                log(f"  {k} -> {v}")

    return mapping

# ==================== 5. 照片重命名 ====================
def rename_photos_in_place(extract_root):
    count = 0
    MAX_PATH = 250  # 留出一些余量
    for root, dirs, files in os.walk(extract_root, topdown=False):
        dir_name = os.path.basename(root)
        for f in files:
            if f.lower().endswith(PHOTO_EXT):
                name, ext = os.path.splitext(f)
                # 正常拼接新文件名
                new_name = f"{name}-{dir_name}{ext}"
                old_path = os.path.join(root, f)
                new_path = os.path.join(root, new_name)

                # 检查新路径长度
                if len(new_path) > MAX_PATH:
                    # 缩短目录名，保留最后30个字符 + "…" 作为标记
                    short_dir = dir_name[-30:] + "…"
                    new_name = f"{name}-{short_dir}{ext}"
                    new_path = os.path.join(root, new_name)
                    log(f"⚠️ 路径过长，已缩短文件名：{new_path}")

                os.rename(old_path, new_path)
                count += 1
    log(f"照片重命名完成，共 {count} 张")


# ==================== 6. 分类并移动（多站点动态目录） ====================
def classify_and_move(extract_root, filename_site_map, date_str, output_base=None):
    moved = 0
    unknown = 0
    if output_base is None:
        output_base = OUTPUT_DIR

    for root, dirs, files in os.walk(extract_root):
        if root == extract_root:
            continue
        dir_name = os.path.basename(root)

        # 判断区域
        target_area = None
        if any(kw in dir_name for kw in DISTRIBUTION_RULES):
            target_area = "配电区"
        elif any(kw in dir_name for kw in SOLAR_RULES):
            target_area = "光伏区"
        if not target_area:
            log(f"⚠️ 跳过未匹配区域的目录：{dir_name}")
            continue

        for f in files:
            if not f.lower().endswith(PHOTO_EXT):
                continue
            src = os.path.join(root, f)
            current_name = f

            # 还原原始文件名（去掉 -目录名 后缀）
            suffix = f"-{dir_name}"
            name_part, ext = os.path.splitext(current_name)
            if name_part.endswith(suffix):
                original_name = name_part[:-len(suffix)] + ext
            else:
                original_name = current_name

            # ========== 增强站点查找（支持歧义检测） ==========
            site = None
            original_lower = original_name.lower()

            # 1. 直接匹配（映射值是列表）
            candidate_sites = filename_site_map.get(original_lower)

            # 2. 尝试去括号匹配（如 IMG_5150(1).JPEG -> IMG_5150.JPEG）
            if not candidate_sites:
                cleaned = re.sub(r'\(\d+\)(?=\.\w+$)', '', original_lower)
                if cleaned != original_lower:
                    candidate_sites = filename_site_map.get(cleaned)

            # 3. 如果仍未找到，尝试模糊匹配（原逻辑保留，用于特殊命名）
            if not candidate_sites:
                name_no_ext = os.path.splitext(original_lower)[0]
                # 去括号后的基本名
                base_clean = os.path.splitext(re.sub(r'\(\d+\)(?=\.\w+$)', '', original_lower))[0]
                sites_found = set()
                for k, v in filename_site_map.items():
                    if base_clean in os.path.splitext(k)[0]:
                        sites_found.update(v)
                if len(sites_found) == 1:
                    candidate_sites = list(sites_found)
                elif len(sites_found) > 1:
                    log(f"⚠️ 多站点歧义（模糊匹配）：{original_name} 可匹配多个站点 {sites_found}")
                    candidate_sites = None

            # 4. 根据候选站点决定最终站点
            if candidate_sites:
                if len(candidate_sites) == 1:
                    site = candidate_sites[0]
                else:
                    # 多站点歧义
                    log(f"⚠️ 多站点歧义：{original_name} 可匹配多个站点 {candidate_sites}，放入未知站点")
                    site = None

            if not site:
                if candidate_sites is None and not filename_site_map.get(original_lower):
                    log(f"⚠️ 未找到照片对应的站点：{original_name}")
                site = "未知站点"
                unknown += 1
            # ====================================================

            # 目标目录
            dest_dir = os.path.join(output_base, f"{site}-{date_str}", target_area)
            os.makedirs(dest_dir, exist_ok=True)

            dst = os.path.join(dest_dir, current_name)
            MAX_PATH = 250
            # 检查目标路径长度
            if len(dst) > MAX_PATH:
                # 缩短照片名，保留原名称（不带目录后缀）并截断
                base_name, ext = os.path.splitext(current_name)
                # 保留不超过200字符的基础名，加上扩展名
                short_base = base_name[:200] + "…"
                current_name = short_base + ext
                dst = os.path.join(dest_dir, current_name)
                log(f"⚠️ 目标路径过长，已缩短目标文件名：{dst}")

            # 同名冲突处理（原逻辑不变）
            if os.path.exists(dst):
                base, ext = os.path.splitext(current_name)
                counter = 1
                while True:
                    new_name = f"{base}_({counter}){ext}"
                    dst = os.path.join(dest_dir, new_name)
                    if not os.path.exists(dst):
                        break
                    counter += 1
            shutil.move(src, dst)
            moved += 1

    log(f"分类移动完成，共处理 {moved} 张照片（其中 {unknown} 张未匹配站点/多站点歧义）")

# # ==================== 主流程-老的单独执行其中一个zip和excel文件====================
# def main():
#     check_directories(cfg)   # 现在 log 已定义，可以安全调用
#     log("======== 每日归档处理开始 ========")
#     extract_root = None
#     try:
#         # 1. 解压
#         zip_path = find_latest_zip()
#         if not zip_path or not os.path.exists(zip_path):
#             log("❌ 未找到任何压缩包，程序终止")
#             return
#
#         extract_root = unzip_file(zip_path)
#
#         # 2. 获取站点前缀和文件修改日期
#         site_prefix, date_str = get_site_and_date(zip_path)
#         # 动态创建本次输出子目录，例如 output/YH光伏巡检-照片巡检-20260519
#         global OUTPUT_DIR
#         OUTPUT_DIR = os.path.join(OUTPUT_DIR, f"{site_prefix}-{date_str}")
#         os.makedirs(OUTPUT_DIR, exist_ok=True)
#         log(f"本次输出目录：{OUTPUT_DIR}")
#
#         # 3. 查找配套 Excel
#         excel_path = find_excel_file(zip_path)
#         if not excel_path:
#             log("❌ 缺少配套 Excel，无法按站点分类，程序终止")
#             return
#
#         # 4. 构建映射
#         site_map = build_filename_site_map(excel_path)
#
#         # 5. 重命名照片
#         rename_photos_in_place(extract_root)
#
#         # 6. 按站点和区域分类移动
#         classify_and_move(extract_root, site_map, date_str)
#
#         log("======== 处理全部完成 ========")
#     except Exception as e:
#         log(f"❌ 发生错误：{e}")
#         raise
#     finally:
#         # 根据配置决定是否清理临时目录（无论成功或失败）
#         if CLEAN_TEMP and extract_root and os.path.exists(extract_root):
#             shutil.rmtree(extract_root)
#             log(f"已自动清理临时解压目录：{extract_root}")

# ==================== 主流程-新的循环处理：遍历每个压缩包，独立执行解压→分类→输出，并将该次生成的子目录移动至最终目录。====================
def main():
    check_directories(cfg)
    log("======== 每日归档处理开始 ========")

    zip_list = find_all_zips()
    if not zip_list:
        log("❌ 未找到任何压缩包，程序终止")
        return

    for zip_path in zip_list:
        log(f"开始处理：{os.path.basename(zip_path)}")
        extract_root = None
        try:
            # 1. 解压
            extract_root = unzip_file(zip_path)

            # 2. 获取站点前缀和文件日期
            site_prefix, date_str = get_site_and_date(zip_path)

            # 3. 动态创建本次输出子目录
            global OUTPUT_DIR
            current_output = os.path.join(OUTPUT_DIR, f"{site_prefix}-{date_str}")
            os.makedirs(current_output, exist_ok=True)
            log(f"本次输出目录：{current_output}")

            # 4. 查找配套 Excel
            excel_path = find_excel_file(zip_path)
            if not excel_path:
                log(f"❌ 缺少配套 Excel，跳过 {os.path.basename(zip_path)}")
                continue

            # 5. 构建映射
            site_map = build_filename_site_map(excel_path)

            # 6. 重命名照片
            rename_photos_in_place(extract_root)

            # 7. 分类移动（使用当前输出子目录）
            classify_and_move(extract_root, site_map, date_str, current_output)

            log(f"FINAL_DIR 原始值：'{FINAL_DIR}'")
            # 8. 搬运至最终目录（移动）
            if FINAL_DIR:
                final_path = os.path.join(FINAL_DIR, f"{site_prefix}-{date_str}")
                # 若目标已存在（可能多日同名），则合并或覆盖
                if os.path.exists(final_path):
                    log(f"最终目录已存在，将合并覆盖：{final_path}")
                shutil.move(current_output, final_path)
                log(f"已移动输出至最终目录：{final_path}")
                # 注意：移动后 current_output 已不存在，无需额外清理
            else:
                log("未配置 FINAL_DIR，输出保留在中间目录")

            log(f"处理完成：{os.path.basename(zip_path)}")

        except Exception as e:
            log(f"❌ 处理 {os.path.basename(zip_path)} 时发生错误：{e}")
            continue  # 继续处理下一个
        finally:
            if CLEAN_TEMP:
                # 删除整个临时工作目录（包含复制的zip和解压内容）
                if os.path.exists(WORK_DIR):
                    shutil.rmtree(WORK_DIR)
                    os.makedirs(WORK_DIR, exist_ok=True)  # 重建空目录，方便后续使用
                    log(f"已清理临时工作目录：{WORK_DIR}")

    # ---- 邮件通知 ----
    if ENABLE_MAIL and MAIL_TO:
        # 构建邮件正文
        body_lines = ["今日巡检照片处理报告："]
        body_lines.extend(zip_summaries)   # zip_summaries 已在循环中填充
        body = "\n".join(body_lines)
        subject = MAIL_SUBJECT_SUCCESS if overall_success else MAIL_SUBJECT_FAIL

        success, msg = send_mail(
            SMTP_SERVER, SMTP_PORT,
            SMTP_USER, SMTP_PASSWORD,
            MAIL_TO,
            subject, body
        )
        if success:
            log("邮件通知发送成功")
        else:
            log(f"邮件通知发送失败：{msg}")
    else:
        log("邮件通知未启用或配置不完整")

    log("======== 所有压缩包处理完成 ========")


if __name__ == "__main__":
    # 测试友好：默认立即执行；若配置了 SCHEDULE_TIME 且未使用 --now 参数，则等待到定时时刻
    if "--now" in sys.argv:
        log("检测到 --now 参数，强制立即执行")
    elif SCHEDULE_TIME:
        try:
            wait_until_scheduled_time(SCHEDULE_TIME)
        except ValueError as e:
            log(f"定时配置错误：{e}，将立即执行")
    main()
