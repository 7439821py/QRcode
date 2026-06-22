import os
import glob
import pandas as pd
from datetime import datetime
from sqlalchemy import create_engine

# ==================== 配置区 ====================
CSV_DIR = r"C:\Users\Administrator\Desktop\聚水潭\2026-06-11"


# MySQL 连接信息
MYSQL_CONFIG = {
    "host": "xxxxxxxxxxx",
    "user": "xxxxxx",
    "password": "xxxxxx",
    "database": "xxxxxx",
    "port": 3306,
    "charset": "utf8mb4"
}

# 关键词 -> 目标表名（包含关系）
# 只要 CSV 文件名中包含这些关键词（任一），就导入对应的表
# 注意顺序：更具体的关键词应放在前面，避免误匹配
KEYWORD_TABLE_MAPPING = [
    ("新媒体成本", "cost_new_media"),
    ("三网成本",   "cost_three_network"),
    ("公司成本",   "cost_company"),
    ("分销成本",   "cost_distribution"),
]

# 手动指定本次导入的时间范围（运行前修改这里）
TIME_RANGE = "2026-06-11"

# 导入模式
IF_EXISTS = 'append'
DB_CHUNK_SIZE = 10000
# ================================================

def get_table_name_from_filename(filename: str) -> str:
    """根据文件名包含的关键词返回目标表名，未匹配则返回 None"""
    for keyword, table_name in KEYWORD_TABLE_MAPPING:
        if keyword in filename:
            return table_name
    return None


def import_csv_to_mysql(csv_path, table_name, engine, if_exists='append'):
    if not os.path.exists(csv_path):
        print(f"❌ 文件不存在: {csv_path}")
        return False

    print(f"\n📄 处理文件: {os.path.basename(csv_path)} -> 表 {table_name}")
    try:
        # 读取表头以确定哪些列存在
        sample = pd.read_csv(csv_path, nrows=0, encoding='utf-8-sig')
        # 需要指定为字符串的列（仅针对存在的列）
        mixed_type_columns = [
            'shop_style_code', 'offline_customer_name', 'product_gb_code',
            'product_main_location', 'order_type', 'shop_product_code'
        ]
        dtype_dict = {col: str for col in mixed_type_columns if col in sample.columns}

        reader = pd.read_csv(csv_path, chunksize=DB_CHUNK_SIZE, encoding='utf-8-sig',
                             dtype=dtype_dict, low_memory=False)
        first_chunk = True
        total_rows = 0
        now = datetime.now()
        for chunk in reader:
            chunk = chunk.replace(['', 'None'], pd.NA)
            chunk['insert_time'] = now
            chunk['time_range'] = TIME_RANGE
            if first_chunk:
                chunk.to_sql(table_name, engine, if_exists=if_exists, index=False)
                first_chunk = False
            else:
                chunk.to_sql(table_name, engine, if_exists='append', index=False)
            total_rows += len(chunk)
            print(f"  已写入 {len(chunk)} 行，累计 {total_rows} 行")
        print(f"✅ 导入完成，共 {total_rows} 行")
        return True
    except Exception as e:
        print(f"❌ 导入失败: {e}")
        return False

def main():
    db_url = f"mysql+pymysql://{MYSQL_CONFIG['user']}:{MYSQL_CONFIG['password']}@{MYSQL_CONFIG['host']}:{MYSQL_CONFIG['port']}/{MYSQL_CONFIG['database']}?charset={MYSQL_CONFIG['charset']}"
    engine = create_engine(db_url)

    # 递归查找文件夹内所有 CSV 文件（含子文件夹）
    all_csv_paths = glob.glob(os.path.join(CSV_DIR, "**", "*.csv"), recursive=True)
    if not all_csv_paths:
        print(f"❌ 在 {CSV_DIR} 中未找到任何 CSV 文件")
        return

    processed = 0
    for csv_path in all_csv_paths:
        filename = os.path.basename(csv_path)
        table_name = get_table_name_from_filename(filename)
        if not table_name:
            print(f"⚠️ 跳过未匹配的文件: {filename}")
            continue
        import_csv_to_mysql(csv_path, table_name, engine, if_exists=IF_EXISTS)
        processed += 1

    print(f"\n🎉 处理完成，共匹配并导入 {processed} 个文件")

if __name__ == "__main__":
    main()
