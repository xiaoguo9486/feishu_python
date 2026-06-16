"""
unpack_photos.py
功能：自动解压 photos/ 下的飞书附件 zip，将所有图片平铺到 photos/ 根目录。
"""
import os
import zipfile
import shutil
import glob

# ---------- 配置 ----------
PHOTO_FOLDER = "D:/exercise/python/photos/"                  # 照片存放文件夹
ZIP_PATTERN = "*巡检*附件*.zip"            # 匹配飞书下载的压缩包名
IMAGE_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.webp'}

def main():
    # 1. 查找压缩包
    search_path = os.path.join(PHOTO_FOLDER, ZIP_PATTERN)
    zip_files = glob.glob(search_path)
    if not zip_files:
        print("❌ 未找到匹配的压缩包，请检查 photos/ 目录")
        return

    print(f"📦 找到 {len(zip_files)} 个压缩包：")
    for zf in zip_files:
        print(f"   - {os.path.basename(zf)}")

    # 2. 解压并移动图片
    for zip_path in zip_files:
        print(f"\n⏳ 正在处理：{os.path.basename(zip_path)}")
        # 解压到临时文件夹
        temp_dir = os.path.join(PHOTO_FOLDER, "_temp_unpack")
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
        os.makedirs(temp_dir, exist_ok=True)

        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        # 3. 遍历所有子文件夹，移动图片到 photos/ 根目录
        moved_count = 0
        for root, dirs, files in os.walk(temp_dir):
            for file in files:
                ext = os.path.splitext(file)[1].lower()
                if ext in IMAGE_EXTENSIONS:
                    src = os.path.join(root, file)
                    dst = os.path.join(PHOTO_FOLDER, file)

                    # 处理重名：若目标已存在，加数字后缀
                    if os.path.exists(dst):
                        base, ext_orig = os.path.splitext(file)
                        counter = 1
                        while True:
                            new_name = f"{base}_{counter}{ext_orig}"
                            dst = os.path.join(PHOTO_FOLDER, new_name)
                            if not os.path.exists(dst):
                                break
                            counter += 1
                        print(f"   ⚠️  重名处理：{file} → {new_name}")

                    shutil.move(src, dst)
                    moved_count += 1

        # 4. 删除临时文件夹
        shutil.rmtree(temp_dir)
        print(f"   ✅ 完成，移动了 {moved_count} 张图片")

        # 5. 可选：删除已处理的压缩包（取消下行注释即删除）
        # os.remove(zip_path)
        # print(f"   🗑️  已删除压缩包：{os.path.basename(zip_path)}")

    print("\n🎉 所有压缩包处理完毕！现在可以运行 generate_report.py 了。")

if __name__ == "__main__":
    main()