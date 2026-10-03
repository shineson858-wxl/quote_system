# coolers.py
# 用途：将 NAKO 工业冷风机报价 Excel 中的整机与配件价格导入 MySQL（solar_quotation_db），
#       并生成阶梯价格矩阵（大型/中型经销商、中型/小型安装商 × 越南本地仓/中国出港）。
# Purpose: Import NAKO industrial evaporative cooler products & accessories from the Excel
#          quotation sheet into MySQL, and build the tiered price matrix by region & trade term.
import pandas as pd
import pymysql
import re
import os

DB_CONFIG = {
    'host': '127.0.0.1',
    'port': 3306,
    'user': 'root',
    'password': '',
    'database': 'solar_quotation_db',
    'charset': 'utf8mb4',
    'autocommit': False
}

def clean_val(val, default=0.0):
    try:
        if pd.isna(val): return default
        return float(val)
    except:
        return default

def find_cooler_excel():
    """自動在目前目錄及同級搜尋冷風機 Excel 檔案"""
    target_names = [
        '工业冷风机报价单 BÁO GIÁ MÁY LÀM MÁT CÔNG NGHIỆP.xlsx',
        'BÁO GIÁ MÁY LÀM MÁT CÔNG NGHIỆP.xlsx',
        '工业冷风机报价单.xlsx'
    ]
    # 1. 優先精準匹配
    for name in target_names:
        if os.path.exists(name):
            return name

    # 2. 模糊掃描目前資料夾
    files = os.listdir('.')
    for f in files:
        if f.endswith(('.xlsx', '.xls')) and ('冷风机' in f or 'MÁY LÀM MÁT' in f or 'LÀM MÁT' in f.upper()):
            return f

    return None

def import_cooler_data():
    excel_file = find_cooler_excel()
    if not excel_file:
        print("\n❌ 錯誤：在目前目錄 D:\\quote_system\\ 下未找到冷風機 Excel 檔案！")
        print("目前目錄下的所有檔案清單如下，請確認檔案是否已放入：")
        for f in os.listdir('.'):
            if f.endswith(('.xlsx', '.xls')):
                print(f"  📄 {f}")
        return

    print(f" 成功找到冷風機報價檔案: {excel_file}，開始解析...")
    
    conn = pymysql.connect(**DB_CONFIG)
    cursor = conn.cursor()
    
    # 讀入原始表格
    df = pd.read_excel(excel_file, header=None)
    
    usd_rate = 25967.0  # 鎖定 VND 換算美金匯率
    items_parsed = []
    
    current_model_code = ""
    current_specs = ""

    # 從第 2 行開始逐行掃描（跳過說明行與列頭）
    for r in range(2, len(df)):
        row = df.iloc[r]
        col_name = str(row[1]).strip() if pd.notna(row[1]) else ""
        
        # 遇到新主機組（解決合併單元格向下填充）
        if col_name and col_name != 'nan':
            lines = [line.strip() for line in col_name.split('\n') if line.strip()]
            m = re.search(r'(HN-[\w]+|TH-[\w]+|NK-[\w]+)', col_name)
            current_model_code = m.group(1) if m else lines[0].replace(' ', '_')
            current_specs = " | ".join(lines[1:])
        
        sub_model = str(row[2]).strip() if pd.notna(row[2]) else ""
        if not sub_model or sub_model == 'nan':
            continue
            
        unit = str(row[3]).strip().replace('\n', '') if pd.notna(row[3]) else "台"
        price_dealer_vnd = clean_val(row[4])   # 經銷價 (VND)
        price_retail_vnd = clean_val(row[5])   # 零售價 (VND)
        
        if price_dealer_vnd <= 0:
            continue
            
        price_dealer_usd = round(price_dealer_vnd / usd_rate, 2)
        price_retail_usd = round(price_retail_vnd / usd_rate, 2)
        
        # 區分是冷風機整機還是專用配件
        is_part = any(sub_model.startswith(k) for k in ['供水泵', '排水泵', '机械浮球阀'])
        clean_sub = re.sub(r'[^\w\s-]', '', sub_model).replace(' ', '_')[:25]
        
        if is_part:
            cat = "冷风机配件"
            full_model = f"NAKO {current_model_code} 配件: {sub_model}"
            spec_text = sub_model
            sku = f"PART-{current_model_code}-{clean_sub}"
        else:
            cat = "工业冷风机"
            full_model = f"NAKO {current_model_code} ({sub_model})"
            spec_text = current_specs
            sku = f"COOLER-{current_model_code}-{clean_sub}"
            
        items_parsed.append({
            "sku": sku,
            "category": cat,
            "model": full_model,
            "spec": spec_text,
            "unit": unit,
            "price_dealer_usd": price_dealer_usd,
            "price_retail_usd": price_retail_usd
        })

    print(f">>> 成功提取 {len(items_parsed)} 項冷風機及配件，正在寫入 MySQL 資料庫...")

    for item in items_parsed:
        # 1. 寫入產品主表
        sql_prod = """
        INSERT INTO tb_product (sku, brand, category, model, spec, pcs_per_pallet, pcs_per_container, warranty_years)
        VALUES (%s, 'NAKO', %s, %s, %s, 1, 20, 2)
        ON DUPLICATE KEY UPDATE category=VALUES(category), spec=VALUES(spec);
        """
        cursor.execute(sql_prod, (item['sku'], item['category'], item['model'], item['spec']))
        cursor.execute("SELECT id FROM tb_product WHERE sku = %s", (item['sku'],))
        prod_id = cursor.fetchone()[0]

        # 2. 寫入階梯價（大型經銷商=經銷底價*0.95，中型經銷商=經銷價，中型安裝商=+5%，小型安裝商=零售價）
        tiers = [
            ("大型经销商", round(item['price_dealer_usd'] * 0.95, 2)),
            ("中型经销商", item['price_dealer_usd']),
            ("中型安装商", round(item['price_dealer_usd'] * 1.05, 2)),
            ("小型安装商", item['price_retail_usd'])
        ]
        
        for region in ['越南本地仓', '中國出港']:
            for tier_name, price in tiers:
                sql_price = """
                INSERT INTO tb_price_matrix (product_id, region, trade_term, customer_tier, currency, unit_price)
                VALUES (%s, %s, 'DAP', %s, 'USD', %s)
                ON DUPLICATE KEY UPDATE unit_price=VALUES(unit_price);
                """
                cursor.execute(sql_price, (prod_id, region, tier_name, price))

    conn.commit()
    cursor.close()
    conn.close()
    print(f"\n🎉 全部 {len(items_parsed)} 款 NAKO 工業冷風機及配件已全部成功入庫！")

if __name__ == '__main__':
    import_cooler_data()