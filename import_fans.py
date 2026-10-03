# import_fans.py
import pandas as pd
import pymysql
import re

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

def import_fan_data():
    conn = pymysql.connect(**DB_CONFIG)
    cursor = conn.cursor()
    excel_file = '2026 一级经销合作价格单.xlsx'
    df_all = pd.read_excel(excel_file, header=None)

    # 基準匯率換算 (以 1 USD = 6.70 CNY, 1 CNY = 3874 VND 換算基準美金與越南盾底價)
    usd_cny_rate = 6.70

    fans_list = []

    # 1. 解析表格上半段 (10 款大型工業吊扇: 4.2M ~ 7.3M)
    for c in range(1, 11):
        price_cny = clean_val(df_all.iloc[1, c])
        if price_cny <= 0: continue
        model = str(df_all.iloc[2, c]).strip()
        diameter = str(df_all.iloc[3, c]).strip()
        blades = str(df_all.iloc[4, c]).strip()
        power = str(df_all.iloc[7, c]).strip()
        air_vol = str(df_all.iloc[8, c]).strip()
        voltage = str(df_all.iloc[9, c]).strip()
        
        sku = f"FAN-{model}-{diameter.replace(' ', '')}-{blades}B"
        full_spec = f"直徑:{diameter} | {blades}葉 | 功率:{power}kW | 風量:{air_vol}m³/h | {voltage}V"
        fans_list.append({
            "sku": sku,
            "brand": "工業節能大吊扇",
            "model": f"{model} ({diameter} {blades}葉)",
            "spec": full_spec,
            "price_cny": price_cny,
            "price_usd": round(price_cny / usd_cny_rate, 2)
        })

    # 2. 解析表格下半段 (5 款中小型工業商業吊扇: 2.4M ~ 3.8M)
    for c in range(1, 6):
        price_cny = clean_val(df_all.iloc[15, c])
        if price_cny <= 0: continue
        model = str(df_all.iloc[16, c]).strip()
        diameter = str(df_all.iloc[17, c]).strip()
        blades = str(df_all.iloc[18, c]).strip()
        power = str(df_all.iloc[21, c]).strip()
        air_vol = str(df_all.iloc[22, c]).strip()
        voltage = str(df_all.iloc[23, c]).strip()
        
        sku = f"FAN-{model}-{diameter.replace(' ', '')}-{blades}B"
        full_spec = f"直徑:{diameter} | {blades}葉 | 功率:{power}kW | 風量:{air_vol}m³/h | {voltage}V"
        fans_list.append({
            "sku": sku,
            "brand": "工業商業吊扇",
            "model": f"{model} ({diameter} {blades}葉)",
            "spec": full_spec,
            "price_cny": price_cny,
            "price_usd": round(price_cny / usd_cny_rate, 2)
        })

    print(f">>> 成功提取 {len(fans_list)} 款工業大吊扇，正在寫入 MySQL 資料庫...")

    for f in fans_list:
        # 寫入產品主表
        sql_prod = """
        INSERT INTO tb_product (sku, brand, category, model, spec, pcs_per_pallet, pcs_per_container, warranty_years)
        VALUES (%s, %s, '工業吊扇通風', %s, %s, 1, 50, 3)
        ON DUPLICATE KEY UPDATE brand=VALUES(brand), spec=VALUES(spec);
        """
        cursor.execute(sql_prod, (f['sku'], f['brand'], f['model'], f['spec']))
        cursor.execute("SELECT id FROM tb_product WHERE sku = %s", (f['sku'],))
        prod_id = cursor.fetchone()[0]

        # 寫入階梯價格矩陣 (以出廠人民幣/美金基準，涵蓋中國出港FOB與越南本地倉DAP)
        # 大型經銷商: 原價; 中型安裝商: 浮動+8%; 小型安裝商/散客: +15%
        tiers = [
            ("大型經銷商", f['price_usd']),
            ("中型經銷商", round(f['price_usd'] * 1.05, 2)),
            ("中型安裝商", round(f['price_usd'] * 1.08, 2)),
            ("小型安裝商", round(f['price_usd'] * 1.15, 2))
        ]
        for region, term in [('中國出港', 'FOB'), ('越南本地倉', 'DAP')]:
            for tier_name, price in tiers:
                sql_price = """
                INSERT INTO tb_price_matrix (product_id, region, trade_term, customer_tier, currency, unit_price)
                VALUES (%s, %s, %s, %s, 'USD', %s)
                ON DUPLICATE KEY UPDATE unit_price=VALUES(unit_price);
                """
                cursor.execute(sql_price, (prod_id, region, term, tier_name, price))

    conn.commit()
    cursor.close()
    conn.close()
    print(" 全部 15 款工業節能吊扇物料及階梯價格已成功入庫！")

if __name__ == '__main__':
    import_fan_data()