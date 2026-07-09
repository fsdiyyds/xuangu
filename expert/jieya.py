import os
import zipfile

# ===================== 请在这里修改你的文件夹路径 =====================
target_folder = r"D:\2026年工作\17 摩洛哥小刚构\摩纳哥小刚构\4标同类型桥梁图纸参考-参考这个文件的图纸\FusionLive_1760702576366"  # Windows 示例
# target_folder = "/home/user/zipfiles"  # Linux/macOS 示例
# ======================================================================

# 遍历文件夹里所有文件
for filename in os.listdir(target_folder):
    # 只处理 .zip 文件（不区分大小写）
    if filename.lower().endswith(".zip"):
        zip_path = os.path.join(target_folder, filename)

        try:
            # 打开 zip 并解压到当前文件夹（覆盖模式）
            with zipfile.ZipFile(zip_path, "r") as zip_ref:
                zip_ref.extractall(target_folder)
            print(f"✅ 成功解压：{filename}")

        except Exception as e:
            print(f"❌ 解压失败：{filename}，原因：{str(e)}")

print("\n🎉 所有 zip 解压完成！")