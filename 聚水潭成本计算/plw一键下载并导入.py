"""
一键运行：下载 → 解压/转换 → 合并 → 导入 MySQL
不修改 plw下载数据.py 和 压缩包合并导入MySQL.py 原文件

使用方法（命令行或 PyCharm Parameters）：
  python plw一键下载并导入.py        全流程：下载 → 解压/转换 → 入库
  python plw一键下载并导入.py 1       同上（默认）
  python plw一键下载并导入.py 2       跳过下载，从解压/转换开始
  python plw一键下载并导入.py 3       跳过下载和解压，只重新导入已有的 CSV
"""

import os
import sys
import subprocess
import glob
import pandas as pd
from datetime import datetime
from sqlalchemy import create_engine

# ── 导入已有的解压/合并/入库函数 ──
sys.path.insert(0, os.path.dirname(__file__))
import 压缩包合并导入MySQL as imp

# ── 配置 ──
MYSQL_CONFIG = imp.MYSQL_CONFIG
FIELD_MAPPING = imp.FIELD_MAPPING
IF_EXISTS = imp.IF_EXISTS
# 让 TIME_RANGE 自动取今天日期（下载当天）
TIME_RANGE = datetime.now().strftime('%Y-%m-%d')
DB_CHUNK_SIZE = imp.DB_CHUNK_SIZE

today = TIME_RANGE
ROOT_FOLDER = fr'C:\Users\Administrator\Desktop\聚水潭\{today}'


def process_single_xlsx(xlsx_path, engine):
    """处理单个 .xlsx 文件（不走解压，直接转 CSV 后导入）"""
    filename = os.path.splitext(os.path.basename(xlsx_path))[0]
    print(f"\n📄 处理单个 Excel: {filename}")

    # 根据文件名匹配目标表
    target_table = imp.根据文件名获取表名(filename)
    if target_table is None:
        print(f"  ⚠️ 文件名「{filename}」未匹配到目标表，跳过")
        return

    # 转 CSV（映射中文字段为英文字段）
    csv_path = xlsx_path.replace('.xlsx', '.csv')
    rows = imp.excel转csv并映射字段(xlsx_path, csv_path, FIELD_MAPPING, append_data=False)
    if rows == 0:
        print(f"  ⚠️ 无数据行，跳过")
        os.remove(csv_path)
        return
    print(f"  ✅ CSV 生成: {os.path.basename(csv_path)} (共 {rows} 行)")

    # 导入 MySQL
    imp.csv导入mysql(csv_path, target_table, engine, IF_EXISTS)

    # 清理临时 CSV
    os.remove(csv_path)


def 解析起始步骤():
    """解析命令行参数，返回起始步骤（1/2/3）"""
    if len(sys.argv) >= 2 and sys.argv[1].isdigit():
        step = int(sys.argv[1])
        if step in (1, 2, 3):
            return step
        print(f"⚠️ 无效步骤: {step}，可选 1/2/3，默认全流程")
        sys.exit(1)
    return 1


def main():
    start_step = 解析起始步骤()

    # ══════════ 步骤1：下载 ══════════
    if start_step <= 1:
        print("=" * 50)
        print("步骤1：下载文件")
        print("=" * 50)
        result = subprocess.run(
            [sys.executable, 'plw下载数据.py'],
            cwd=os.path.dirname(__file__),
        )
        if result.returncode != 0:
            print(f"\n❌ plw下载数据.py 执行失败（退出码 {result.returncode}），终止")
            sys.exit(1)
    else:
        print("⏭️  跳过步骤1：下载")

    # ── 检查下载目录 ──
    if not os.path.exists(ROOT_FOLDER):
        print(f"\n❌ 目录不存在: {ROOT_FOLDER}")
        print("  请先执行下载，或指定正确日期")
        sys.exit(1)

    # ── 连接数据库 ──
    db_url = (
        f"mysql+pymysql://{MYSQL_CONFIG['user']}:{MYSQL_CONFIG['password']}"
        f"@{MYSQL_CONFIG['host']}:{MYSQL_CONFIG['port']}"
        f"/{MYSQL_CONFIG['database']}?charset={MYSQL_CONFIG['charset']}"
    )
    engine = create_engine(db_url)

    # ══════════ 步骤2：解压/转换 → 合并成 CSV ══════════
    if start_step <= 2:
        # ── 处理 .zip 文件（解压 → 合并为 CSV → 导入） ──
        zip_files = glob.glob(os.path.join(ROOT_FOLDER, "*.zip"))
        if zip_files:
            print(f"\n📁 找到 {len(zip_files)} 个压缩包")
            for zip_path in zip_files:
                try:
                    imp.处理压缩包(zip_path, ROOT_FOLDER, engine,
                                    FIELD_MAPPING, IF_EXISTS)
                except Exception as e:
                    print(f"  ❌ 处理失败: {e}")
                    import traceback
                    traceback.print_exc()
        else:
            print("\n📁 无压缩包")

        # ── 处理 .xlsx 文件（转 CSV → 导入） ──
        xlsx_files = glob.glob(os.path.join(ROOT_FOLDER, "*.xlsx"))
        if xlsx_files:
            print(f"\n📁 找到 {len(xlsx_files)} 个 Excel 文件")
            for xlsx_path in xlsx_files:
                try:
                    process_single_xlsx(xlsx_path, engine)
                except Exception as e:
                    print(f"  ❌ 处理失败: {e}")
                    import traceback
                    traceback.print_exc()
        else:
            print("📁 无 Excel 文件")
    else:
        print("⏭️  跳过步骤2：解压/转换")

    # ══════════ 步骤3：重新导入已有的 CSV ══════════
    if start_step <= 2:
        # 上一步已经导入了，不需要重复导入
        pass
    else:
        print("=" * 50)
        print("步骤3：重新导入已有的 CSV 文件")
        print("=" * 50)
        # 扫描 ROOT_FOLDER 下所有子目录中的 CSV（zip 解压产物）
        csv_files = glob.glob(os.path.join(ROOT_FOLDER, "**", "*.csv"), recursive=True)
        if not csv_files:
            print("📁 未找到 CSV 文件")
        for csv_path in csv_files:
            filename = os.path.splitext(os.path.basename(csv_path))[0]
            target_table = imp.根据文件名获取表名(filename)
            if target_table is None:
                print(f"  ⚠️ {filename} 未匹配到目标表，跳过")
                continue
            print(f"\n📄 导入: {os.path.basename(csv_path)} → {target_table}")
            try:
                imp.csv导入mysql(csv_path, target_table, engine, IF_EXISTS)
            except Exception as e:
                print(f"  ❌ 导入失败: {e}")
                import traceback
                traceback.print_exc()

    print(f"\n🎉 全部完成！")


if __name__ == "__main__":
    main()
