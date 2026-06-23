import os
import re
import shutil
import datetime

# ---------- 配置区 ----------
SOURCE_DIR = r"E:\光伏运维\01流水资料\自动下载目录"
DEST_DIRS = [
    r"D:\exercise\python\photos",
    r"D:\exercise\python\pthto"
]
TARGET_NAMES = [
    "DH光伏巡检-照片巡检_xg简易自动巡检系统_照片-昨日.xlsx",
    "DH光伏巡检-照片巡检_xg简易自动巡检系统_照片-昨日_附件.zip",
    "YH光伏巡检-照片巡检_xg简易自动巡检系统_照片-昨日.xlsx",
    "YH光伏巡检-照片巡检_xg简易自动巡检系统_照片-昨日_附件.zip"
]

# ---------- 工具函数 ----------
def get_pattern(fullname):
    base, ext = os.path.splitext(fullname)
    # 允许括号前任意空格（0或多个），括号支持半角/全角，数字部分不变
    return re.compile(re.escape(base) + r"(?:[ \u3000]*[\(\（]\d+[\)\）])?" + re.escape(ext) + "$")

def find_file(source_dir, pattern):
    try:
        files = os.listdir(source_dir)
    except Exception as e:
        raise Exception(f"无法列出源目录: {e}")
    matches = [f for f in files if pattern.match(f)]
    if not matches:
        # 调试：打印包含 'YH' 且含 '附件' 的所有文件
        debug_files = [f for f in files if 'YH' in f and '附件' in f]
        if debug_files:
            print(f"[调试] 与 'YH-附件' 相关的文件有: {debug_files}")
        return None
    matches.sort()
    return matches[0]

def copy_file(src_path, dest_dir):
    """复制文件到目标目录，自动创建目标目录"""
    os.makedirs(dest_dir, exist_ok=True)
    dest_path = os.path.join(dest_dir, os.path.basename(src_path))
    shutil.copy2(src_path, dest_path)
    return dest_path

def resave_excel_with_openpyxl(file_paths):
    """使用 openpyxl 加载并保存 Excel（不依赖 MS Excel）"""
    import openpyxl
    success = []
    fail = []
    for fp in file_paths:
        try:
            wb = openpyxl.load_workbook(fp)
            wb.save(fp)
            success.append(fp)
        except Exception as e:
            fail.append((fp, str(e)))
    return success, fail

# ---------- 主流程 ----------
def main():
    # 创建日志目录
    log_dir = r"D:\exercise\python\logs"
    os.makedirs(log_dir, exist_ok=True)

    # 生成带日期的日志文件名
    log_date = datetime.datetime.now().strftime("%Y-%m-%d")
    log_file = os.path.join(log_dir, f"find_file_copy_{log_date}.log")

    log_lines = []
    exec_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_lines.append(f"执行时间: {exec_time}")
    log_lines.append("=" * 60)

    all_excel_copies = []
    errors = []
    warnings = []

    # 1. 查找并复制
    for name in TARGET_NAMES:
        pattern = get_pattern(name)
        try:
            found = find_file(SOURCE_DIR, pattern)
        except Exception as e:
            err_msg = f"访问源目录出错: {e}"
            errors.append(err_msg)
            log_lines.append(f"[错误] {name} -> {err_msg}")
            continue

        if found is None:
            warn_msg = f"未找到匹配文件: {name}"
            warnings.append(warn_msg)
            log_lines.append(f"[警告] {warn_msg}")
            continue

        src_path = os.path.join(SOURCE_DIR, found)
        log_lines.append(f"[找到] {name} -> 源文件: {found}")

        for dest in DEST_DIRS:
            try:
                dest_path = copy_file(src_path, dest)
                log_lines.append(f"  复制成功: {dest_path}")
                if name.lower().endswith(".xlsx"):
                    all_excel_copies.append(dest_path)
            except Exception as e:
                err_msg = f"复制到 {dest} 失败: {e}"
                errors.append(err_msg)
                log_lines.append(f"  [错误] {err_msg}")

    # 2. Excel 重新保存（使用 openpyxl）
    if all_excel_copies:
        log_lines.append("\n--- Excel 文件重新另存 (openpyxl) ---")
        try:
            ok, fail = resave_excel_with_openpyxl(all_excel_copies)
            for fp in ok:
                log_lines.append(f"  另存成功: {fp}")
            for fp, reason in fail:
                log_lines.append(f"  [另存失败] {fp}: {reason}")
                errors.append(f"Excel 另存失败: {fp} - {reason}")
        except Exception as e:
            errors.append(f"Excel 另存操作整体失败: {e}")
            log_lines.append(f"  [严重错误] {e}")

    # 3. 汇总
    log_lines.append("\n" + "=" * 60)
    if warnings:
        log_lines.append(f"存在 {len(warnings)} 条警告（业务上可能正常）:")
        for idx, w in enumerate(warnings, 1):
            log_lines.append(f"  {idx}. [警告] {w}")

    if errors:
        log_lines.append(f"\n本次作业存在 {len(errors)} 个错误:")
        for idx, err in enumerate(errors, 1):
            log_lines.append(f"  {idx}. {err}")
    else:
        log_lines.append("所有操作已成功完成。")

    if not errors and not warnings:
        log_lines.append("所有文件复制及另存操作均已成功完成。")

    log_lines.append(f"\n汇总生成时间: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    summary_text = "\n".join(log_lines)

    # 输出并保存
    print(summary_text)
    with open(log_file, "w", encoding="utf-8") as f:
        f.write(summary_text)
    print(f"\n日志已保存至: {os.path.abspath(log_file)}")

if __name__ == "__main__":
    main()