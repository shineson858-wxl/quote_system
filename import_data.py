# import_data.py
# 用途：从“0513-东南亚销售指导价”Excel 批量导入汇率、逆变器/储能电池及东南亚本地仓组件参考价到 MySQL。
# Purpose: Bulk-import exchange rates, inverters/batteries and Southeast-Asia local-warehouse
#          module reference prices from the Excel price list into MySQL.
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

def main():
    conn = pymysql.connect(**DB_CONFIG)
    cursor = conn.cursor()
    excel_path = '0513-东南亚销售指导价(1).xlsx'
    xls = pd.ExcelFile(excel_path)

    print(">>> 1/3 正在清洗并导入【汇率与基准参数】...")
    df_rate = pd.read_excel(xls, '汇率')
    for _, row in df_rate.iterrows():
        pair = str(row.iloc[0]).strip()
        rate_val = clean_val(row.iloc[1])
        if pair and rate_val > 0 and pair != 'nan':
            clean_key = pair.replace('-', '_').replace(' ', '')
            sql = """
            INSERT INTO tb_exchange_rate (pair, rate) 
            VALUES (%s, %s) 
            ON DUPLICATE KEY UPDATE rate=VALUES(rate);
            """
            cursor.execute(sql, (clean_key, rate_val))

    print(">>> 2/3 正在清洗并导入【逆变器与储能电池】(硕日 / 固德威)...")
    sheets_inverters = ['逆变器-硕日', '逆变器-固德威', '逆变器-10年质保']
    for sheet_name in sheets_inverters:
        df_inv = pd.read_excel(xls, sheet_name, header=None)
        
        # 寻找列头行
        header_idx = -1
        for i in range(min(6, len(df_inv))):
            row_text = "".join([str(x) for x in df_inv.iloc[i].values])
            if '产品型号' in row_text or '物料编码' in row_text:
                header_idx = i
                break
        if header_idx == -1: 
            continue

        warranty = 10 if '10年' in sheet_name else 5
        brand = "硕日" if "硕日" in sheet_name else "固德威"

        for r in range(header_idx + 1, len(df_inv)):
            row = df_inv.iloc[r]
            model = str(row[3]).strip() if pd.notna(row[3]) else ""
            if not model or model == 'nan': 
                continue

            raw_cat = str(row[1]).strip() if pd.notna(row[1]) else ""
            cat = "储能逆变器"
            if "电池" in raw_cat or "Battery" in raw_cat:
                cat = "储能电池"
            elif "电表" in raw_cat:
                cat = "系统辅材"

            raw_sku = str(row[2]).strip() if pd.notna(row[2]) else ""
            sku = raw_sku if raw_sku and raw_sku != 'nan' else f"{brand}-{model}-{warranty}Y"
            spec = str(row[4]).strip() if pd.notna(row[4]) and str(row[4]) != 'nan' else ""
            cont_qty = int(clean_val(row[5]))
            pallet_qty = int(clean_val(row[6]))

            # 插入或更新基础物料
            sql_prod = """
            INSERT INTO tb_product (sku, brand, category, model, spec, pcs_per_pallet, pcs_per_container, warranty_years)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
            ON DUPLICATE KEY UPDATE 
                pcs_per_pallet=VALUES(pcs_per_pallet), 
                pcs_per_container=VALUES(pcs_per_container);
            """
            cursor.execute(sql_prod, (sku, brand, cat, model, spec, pallet_qty, cont_qty, warranty))
            cursor.execute("SELECT id FROM tb_product WHERE sku = %s", (sku,))
            prod_id = cursor.fetchone()[0]

            # 提取 FOB 美元阶梯价：列8~11分别对应 [大型经销商, 中型经销商, 中型安装商, 小型安装商]
            tiers = ["大型经销商", "中型经销商", "中型安装商", "小型安装商"]
            for idx, tier_name in enumerate(tiers):
                col_num = 8 + idx
                if col_num < len(row):
                    price = clean_val(row[col_num])
                    if price > 0:
                        sql_price = """
                        INSERT INTO tb_price_matrix (product_id, region, trade_term, customer_tier, currency, unit_price)
                        VALUES (%s, '中国出港', 'FOB', %s, 'USD', %s)
                        ON DUPLICATE KEY UPDATE unit_price=VALUES(unit_price);
                        """
                        cursor.execute(sql_price, (prod_id, tier_name, price))

    print(">>> 3/3 正在清洗并导入【东南亚本地仓与组件参考价】...")
    df_sea_mod = pd.read_excel(xls, '东南亚仓组件', header=None)
    cur_country = "泰国"
    cur_term = "EXW"

    for r in range(1, len(df_sea_mod)):
        row = df_sea_mod.iloc[r]
        cell_0 = str(row[0]).strip()
        if '越南' in cell_0:
            cur_country = "越南"; cur_term = "DAP"; continue
        elif '马来西亚' in cell_0:
            cur_country = "马来西亚"; cur_term = "DAP"; continue
        elif '国家' in cell_0 or 'SKU' in cell_0:
            continue

        model = str(row[3]).strip() if pd.notna(row[3]) else ""
        if not model or model == 'nan': 
            continue

        raw_sku = str(row[1]).strip() if pd.notna(row[1]) else ""
        sku = raw_sku if raw_sku and raw_sku != 'nan' else f"MOD-{model}"
        
        # 提取功率(W)
        watts = 0.0
        match = re.search(r'(\d{3})', model)
        if match:
            watts = float(match.group(1))

        brand = "爱旭" if "爱旭" in model else ("高景" if "高景" in model else "一线品牌")

        sql_prod = """
        INSERT INTO tb_product (sku, brand, category, model, spec, power_w)
        VALUES (%s, %s, '光伏组件', %s, %s, %s)
        ON DUPLICATE KEY UPDATE power_w=VALUES(power_w);
        """
        cursor.execute(sql_prod, (sku, brand, model, f"{watts}W", watts))
        cursor.execute("SELECT id FROM tb_product WHERE sku = %s", (sku,))
        prod_id = cursor.fetchone()[0]

        # 本地提货阶梯：整柜、整托、散单
        mod_tiers = [("整柜", clean_val(row[4])), ("整托", clean_val(row[5])), ("散单", clean_val(row[6]))]
        for t_name, t_price in mod_tiers:
            if t_price > 0:
                per_w = round(t_price / watts, 4) if watts > 0 else None
                sql_price = """
                INSERT INTO tb_price_matrix (product_id, region, trade_term, customer_tier, currency, unit_price, price_per_w)
                VALUES (%s, %s, %s, %s, 'USD', %s, %s)
                ON DUPLICATE KEY UPDATE unit_price=VALUES(unit_price), price_per_w=VALUES(price_per_w);
                """
                cursor.execute(sql_price, (prod_id, f"{cur_country}本地仓", cur_term, t_name, t_price, per_w))

    conn.commit()
    cursor.close()
    conn.close()
    print("\n 数据清洗与灌库全部完成！可打开 phpMyAdmin 验证。")

if __name__ == '__main__':
    main()