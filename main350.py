# -*- coding: utf-8 -*-
"""
==================================================================================================
項目名稱 (Project): GALAXYTECK 智慧快速報價與客戶關係管理系統 (企業標準網絡版)
版本編號 (Version): v3.5.0 CRM Order Linkage & Document Lifecycle Edition
發布日期 (Date): 2026年10月
軟體授權與智慧財產權聲明 (Intellectual Property & Copyright):
    版權所有 (C) 2024-2026 CÔNG TY TNHH TẬP ĐOÀN CÔNG NGHỆ GALAXY VIỆT NAM
    (GALAXY VIETNAM TECHNOLOGY CORPORATION COMPANY LIMITED / 銀河科技集團(越南)有限公司)
    企業代碼 / Mã số doanh nghiệp: 2301296587
    地址 / Địa chỉ: Số 36 Yết Kiêu, Phường Kinh Bắc, Tỉnh Bắc Ninh, Việt Nam
    未經本公司書面授權許可，嚴禁任何形式之代碼反編譯、篡改、未授權商業複製或對外再分發。
==================================================================================================
"""

from fastapi import FastAPI, HTTPException, UploadFile, File
from fastapi.responses import HTMLResponse
from pydantic import BaseModel
from typing import List, Optional, Any
from decimal import Decimal, ROUND_HALF_UP
import pymysql
import datetime
import uvicorn
import pandas as pd
import io
import json
import uuid
import re

app = FastAPI(
    title="GalaxyTeck 智慧快速報價系統",
    version="3.5.0",
    description="企業級專業報價軟體：具備 CRM 客戶訂單穿透查詢與調出維護、多單據發號器、自定義物料自動入庫、Tailwind JIT、Logo持久化與A4防跑版"
)

# ================= 1. 資料庫連線配置 (XAMPP MySQL) =================
DB_CONFIG = {
    'host': '127.0.0.1',
    'port': 3306,
    'user': 'root',
    'password': '',
    'database': 'solar_quotation_db',
    'charset': 'utf8mb4',
    'cursorclass': pymysql.cursors.DictCursor,
    'autocommit': True
}

def get_db():
    return pymysql.connect(**DB_CONFIG)

def clean_val(val, default=0.0):
    try:
        if pd.isna(val): return default
        return float(val)
    except:
        return default

# ================= 2. 高級全自動自愈引擎 (程式啟動自動檢查修復表結構) =================
@app.on_event("startup")
def self_healing_db():
    try:
        conn = get_db()
        with conn.cursor() as cursor:
            # 1. 確保全局企業設置表
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS `tb_company_setting` (
                    `id` INT PRIMARY KEY DEFAULT 1,
                    `logo_data` LONGTEXT DEFAULT NULL,
                    `company_name_vi` VARCHAR(255) DEFAULT 'CÔNG TY TNHH TẬP ĐOÀN CÔNG NGHỆ GALAXY VIỆT NAM',
                    `company_name_cn` VARCHAR(255) DEFAULT '銀河科技集團（越南）有限公司',
                    `tax_code` VARCHAR(50) DEFAULT '2301296587',
                    `contact_info` VARCHAR(255) DEFAULT 'Điện thoại: 0325.609.218 | Email: galaxyteck6886@gmail.com',
                    `address` VARCHAR(255) DEFAULT 'Số 36 Yết Kiêu, Phường Kinh Bắc, Tỉnh Bắc Ninh, Việt Nam',
                    `bank_info` VARCHAR(255) DEFAULT 'STK (Techcombank): 314431 - Ngân hàng TMCP Kỹ thương Việt Nam',
                    `updated_at` TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
            """)
            cursor.execute("INSERT INTO `tb_company_setting` (`id`, `tax_code`) VALUES (1, '2301296587') ON DUPLICATE KEY UPDATE tax_code=VALUES(tax_code);")

            # 2. 確保多版式範本表
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS `tb_layout_template` (
                    `id` INT AUTO_INCREMENT PRIMARY KEY,
                    `template_name` VARCHAR(100) NOT NULL,
                    `is_default` TINYINT(1) DEFAULT 0,
                    `logo_data` LONGTEXT DEFAULT NULL,
                    `name_vi` VARCHAR(255) DEFAULT 'CÔNG TY TNHH TẬP ĐOÀN CÔNG NGHỆ GALAXY VIỆT NAM',
                    `name_cn` VARCHAR(255) DEFAULT '銀河科技集團（越南）有限公司',
                    `tax_code` VARCHAR(50) DEFAULT '2301296587',
                    `contact_info` VARCHAR(255) DEFAULT 'Điện thoại: 0325.609.218 | Email: galaxyteck6886@gmail.com',
                    `address` VARCHAR(255) DEFAULT 'Số 36 Yết Kiêu, Phường Kinh Bắc, Tỉnh Bắc Ninh, Việt Nam',
                    `bank_info` VARCHAR(255) DEFAULT 'STK (Techcombank): 314431 - Ngân hàng TMCP Kỹ thương Việt Nam',
                    `title_text` VARCHAR(255) DEFAULT 'BẢNG BÁO GIÁ 商業報價單',
                    `greeting_text` TEXT DEFAULT NULL,
                    `lang_mode` VARCHAR(20) DEFAULT 'vi_zh',
                    `template_style` VARCHAR(50) DEFAULT 'standard',
                    `terms_json` LONGTEXT DEFAULT NULL,
                    `created_at` TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    `updated_at` TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
            """)

            # 3. 確保 CRM 客戶管理表
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS `tb_customer` (
                    `id` INT AUTO_INCREMENT PRIMARY KEY,
                    `company_name` VARCHAR(255) NOT NULL,
                    `short_name` VARCHAR(100) DEFAULT '',
                    `tax_code` VARCHAR(50) DEFAULT '',
                    `contact_person` VARCHAR(100) DEFAULT '',
                    `phone` VARCHAR(50) DEFAULT '',
                    `email` VARCHAR(100) DEFAULT '',
                    `address` VARCHAR(255) DEFAULT '',
                    `default_tier` VARCHAR(50) DEFAULT '中型經銷商',
                    `payment_terms` VARCHAR(255) DEFAULT 'TT 50% 預付, 50% 發貨前',
                    `notes` TEXT DEFAULT NULL,
                    `created_at` TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    `updated_at` TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
            """)

            # 4. 確保主報價單表存在且 ID 為自增主鍵
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS `tb_quotation` (
                    `id` INT AUTO_INCREMENT PRIMARY KEY,
                    `quote_no` VARCHAR(100) NOT NULL,
                    `customer_name` VARCHAR(255) NOT NULL,
                    `created_at` TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    `updated_at` TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
            """)
            try:
                cursor.execute("ALTER TABLE `tb_quotation` MODIFY COLUMN `id` INT NOT NULL AUTO_INCREMENT;")
            except Exception:
                pass

            # 補齊 tb_quotation 欄位
            alter_quotation_cols = [
                ("status", "VARCHAR(50) DEFAULT '已生效'"),
                ("layout_template_id", "INT DEFAULT 1"),
                ("customer_id", "INT DEFAULT 0"),
                ("contact_person", "VARCHAR(100) DEFAULT ''"),
                ("phone_email", "VARCHAR(100) DEFAULT ''"),
                ("sales_rep", "VARCHAR(100) DEFAULT 'Thanh Bình 清平'"),
                ("customer_tier", "VARCHAR(50) DEFAULT '中型經銷商'"),
                ("destination", "VARCHAR(100) DEFAULT '越南本地倉 (DAP)'"),
                ("quote_currency", "VARCHAR(20) DEFAULT 'VND'"),
                ("exchange_rate", "DECIMAL(15, 4) DEFAULT 1.0"),
                ("total_base_cost_usd", "DECIMAL(15, 2) DEFAULT 0.0"),
                ("markup_rate", "DECIMAL(8, 2) DEFAULT 0.0"),
                ("tax_rate", "DECIMAL(8, 2) DEFAULT 10.0"),
                ("final_total_amount", "DECIMAL(15, 2) DEFAULT 0.0"),
                ("valid_days", "INT DEFAULT 3"),
                ("terms_text", "TEXT DEFAULT NULL"),
                ("terms_json", "LONGTEXT DEFAULT NULL"),
                ("lang_mode", "VARCHAR(20) DEFAULT 'vi_zh'"),
                ("template_style", "VARCHAR(50) DEFAULT 'standard'"),
                ("custom_title", "VARCHAR(255) DEFAULT 'BẢNG BÁO GIÁ 商業報價單'"),
                ("custom_greeting", "TEXT DEFAULT NULL"),
                ("company_logo_data", "LONGTEXT DEFAULT NULL"),
                ("company_name_vi", "VARCHAR(255) DEFAULT 'CÔNG TY TNHH TẬP ĐOÀN CÔNG NGHỆ GALAXY VIỆT NAM'"),
                ("company_name_cn", "VARCHAR(255) DEFAULT '銀河科技集團（越南）有限公司'"),
                ("company_tax_code", "VARCHAR(50) DEFAULT '2301296587'"),
                ("company_contact_info", "VARCHAR(255) DEFAULT 'Điện thoại: 0325.609.218 | Email: galaxyteck6886@gmail.com'"),
                ("company_address", "VARCHAR(255) DEFAULT 'Số 36 Yết Kiêu, Phường Kinh Bắc, Tỉnh Bắc Ninh, Việt Nam'"),
                ("company_bank_info", "VARCHAR(255) DEFAULT 'STK (Techcombank): 314431 - Ngân hàng TMCP Kỹ thương Việt Nam'")
            ]
            for col, col_type in alter_quotation_cols:
                try:
                    cursor.execute(f"ALTER TABLE `tb_quotation` ADD COLUMN `{col}` {col_type};")
                except Exception:
                    pass

            # 5. 確保報價明細表欄位
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS `tb_quotation_item` (
                    `id` INT AUTO_INCREMENT PRIMARY KEY,
                    `quotation_id` INT NOT NULL,
                    `product_id` INT DEFAULT 0
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
            """)
            try:
                cursor.execute("ALTER TABLE `tb_quotation_item` MODIFY COLUMN `id` INT NOT NULL AUTO_INCREMENT;")
            except Exception:
                pass

            alter_item_cols = [
                ("item_name", "VARCHAR(150) DEFAULT ''"),
                ("item_model", "VARCHAR(150) DEFAULT ''"),
                ("item_spec", "VARCHAR(255) DEFAULT ''"),
                ("unit", "VARCHAR(20) DEFAULT '台 / Bộ'"),
                ("quantity", "INT DEFAULT 1"),
                ("applied_tier", "VARCHAR(50) DEFAULT '中型經銷商'"),
                ("base_price_usd", "DECIMAL(15, 2) DEFAULT 0.00"),
                ("base_price", "DECIMAL(15, 2) DEFAULT 0.00"),
                ("custom_markup", "DECIMAL(8, 2) DEFAULT NULL"),
                ("quote_unit_price", "DECIMAL(15, 2) DEFAULT 0.00"),
                ("subtotal_amount", "DECIMAL(15, 2) DEFAULT 0.00"),
                ("note", "VARCHAR(255) DEFAULT ''")
            ]
            for col, col_type in alter_item_cols:
                try:
                    cursor.execute(f"ALTER TABLE `tb_quotation_item` ADD COLUMN `{col}` {col_type};")
                except Exception:
                    pass

        conn.close()
        print("資料庫自愈巡檢完成：CRM 客戶與訂單鏈路已 100% 準備就緒！")
    except Exception as e:
        print("資料庫巡檢通知:", e)

# ================= 3. 後端單號智慧發號器 =================
@app.get("/api/quotes/generate_next_no")
def generate_next_quote_no():
    now = datetime.datetime.now()
    prefix = f"BG-{now.strftime('%Y.%m.%d')}"
    max_seq = 0
    try:
        conn = get_db()
        with conn.cursor() as cursor:
            cursor.execute("SELECT quote_no FROM tb_quotation WHERE quote_no LIKE %s", (f"{prefix}%",))
            rows = cursor.fetchall()
            for r in rows:
                qno = r.get('quote_no', '')
                m = re.search(r'-(\d+)$', qno)
                if m:
                    seq = int(m.group(1))
                    if seq > max_seq:
                        max_seq = seq
            conn.close()
    except Exception as e:
        print("獲取單號失敗:", e)
    next_no = f"{prefix}-{str(max_seq + 1).zfill(3)}"
    return {"next_quote_no": next_no, "date_str": now.strftime('%d.%m.%Y')}

# ================= 4. 全局企業設定與 Logo 持久化 API =================
class LogoSaveRequest(BaseModel):
    logo_data: Optional[str] = None

@app.get("/api/system/global_settings")
def get_global_settings():
    try:
        conn = get_db()
        with conn.cursor() as cursor:
            cursor.execute("SELECT * FROM tb_company_setting WHERE id = 1 LIMIT 1")
            row = cursor.fetchone()
            conn.close()
            return row or {}
    except Exception as e:
        print("獲取全局設置失敗:", e)
        return {}

@app.post("/api/system/save_global_logo")
def save_global_logo(req: LogoSaveRequest):
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            cursor.execute("""
                INSERT INTO tb_company_setting (id, logo_data)
                VALUES (1, %s)
                ON DUPLICATE KEY UPDATE logo_data = VALUES(logo_data)
            """, (req.logo_data,))
            cursor.execute("""
                UPDATE tb_layout_template SET logo_data = %s WHERE is_default = 1
            """, (req.logo_data,))
        conn.close()
        return {"status": "success", "message": "公司 Logo 已成功永久保存！"}
    except Exception as e:
        if conn: conn.close()
        raise HTTPException(status_code=500, detail=str(e))

# ================= 5. 多版式範本庫 API =================
class LayoutTemplateModel(BaseModel):
    id: Optional[int] = None
    template_name: str
    is_default: Optional[int] = 0
    logo_data: Optional[str] = None
    name_vi: Optional[str] = "CÔNG TY TNHH TẬP ĐOÀN CÔNG NGHỆ GALAXY VIỆT NAM"
    name_cn: Optional[str] = "銀河科技集團（越南）有限公司"
    tax_code: Optional[str] = "2301296587"
    contact_info: Optional[str] = "Điện thoại: 0325.609.218 | Email: galaxyteck6886@gmail.com"
    address: Optional[str] = "Số 36 Yết Kiêu, Phường Kinh Bắc, Tỉnh Bắc Ninh, Việt Nam"
    bank_info: Optional[str] = "STK (Techcombank): 314431 - Ngân hàng TMCP Kỹ thương Việt Nam"
    title_text: Optional[str] = "BẢNG BÁO GIÁ 商業報價單"
    greeting_text: Optional[str] = "Chúng tôi xin gửi đến Quý khách hàng bảng báo giá như sau / 我司現向貴司呈報以下報價表:"
    lang_mode: Optional[str] = "vi_zh"
    template_style: Optional[str] = "standard"
    terms_list: Optional[List[str]] = []

@app.get("/api/layout_templates/list")
def list_layout_templates():
    try:
        conn = get_db()
        with conn.cursor() as cursor:
            cursor.execute("""
                SELECT id, template_name, is_default, logo_data, name_vi, name_cn, 
                       tax_code, contact_info, address, bank_info, title_text, greeting_text, 
                       lang_mode, template_style, terms_json, created_at
                FROM tb_layout_template
                ORDER BY is_default DESC, id ASC
            """)
            rows = cursor.fetchall()
            for r in rows:
                if 'created_at' in r and r['created_at']:
                    r['created_at'] = str(r['created_at'])
                try:
                    r['terms_list'] = json.loads(r['terms_json']) if r['terms_json'] else []
                except:
                    r['terms_list'] = []
            conn.close()
            return rows
    except Exception as e:
        print("獲取版式範本失敗:", e)
        return []

@app.post("/api/layout_templates/save")
def save_layout_template(tpl: LayoutTemplateModel):
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            terms_json = json.dumps(tpl.terms_list, ensure_ascii=False)
            if tpl.is_default == 1:
                cursor.execute("UPDATE tb_layout_template SET is_default = 0")

            if tpl.id and tpl.id > 0:
                cursor.execute("""
                    UPDATE tb_layout_template SET
                        template_name = %s, is_default = %s, logo_data = %s,
                        name_vi = %s, name_cn = %s, tax_code = %s,
                        contact_info = %s, address = %s, bank_info = %s,
                        title_text = %s, greeting_text = %s, lang_mode = %s,
                        template_style = %s, terms_json = %s
                    WHERE id = %s
                """, (
                    tpl.template_name, tpl.is_default, tpl.logo_data,
                    tpl.name_vi, tpl.name_cn, tpl.tax_code,
                    tpl.contact_info, tpl.address, tpl.bank_info,
                    tpl.title_text, tpl.greeting_text, tpl.lang_mode,
                    tpl.template_style, terms_json, tpl.id
                ))
                tid = tpl.id
            else:
                cursor.execute("""
                    INSERT INTO tb_layout_template (
                        template_name, is_default, logo_data, name_vi, name_cn,
                        tax_code, contact_info, address, bank_info,
                        title_text, greeting_text, lang_mode, template_style, terms_json
                    ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """, (
                    tpl.template_name, tpl.is_default, tpl.logo_data,
                    tpl.name_vi, tpl.name_cn, tpl.tax_code,
                    tpl.contact_info, tpl.address, tpl.bank_info,
                    tpl.title_text, tpl.greeting_text, tpl.lang_mode,
                    tpl.template_style, terms_json
                ))
                tid = cursor.lastrowid
            conn.close()
            return {"status": "success", "id": tid, "message": "版式範本保存成功！"}
    except Exception as e:
        if conn: conn.close()
        raise HTTPException(status_code=500, detail=str(e))

@app.delete("/api/layout_templates/{template_id}")
def delete_layout_template(template_id: int):
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            cursor.execute("DELETE FROM tb_layout_template WHERE id = %s", (template_id,))
            conn.close()
            return {"status": "success", "message": "版式範本已刪除！"}
    except Exception as e:
        if conn: conn.close()
        raise HTTPException(status_code=500, detail=str(e))

# ================= 6. CRM 客戶管理 API (含客戶專屬訂單穿透查詢) =================
class CustomerModel(BaseModel):
    id: Optional[int] = None
    company_name: str
    short_name: Optional[str] = ""
    tax_code: Optional[str] = ""
    contact_person: Optional[str] = ""
    phone: Optional[str] = ""
    email: Optional[str] = ""
    address: Optional[str] = ""
    default_tier: Optional[str] = "中型經銷商"
    payment_terms: Optional[str] = "TT 50% 預付, 50% 發貨前"
    notes: Optional[str] = ""

@app.get("/api/customers/list")
def list_customers():
    try:
        conn = get_db()
        with conn.cursor() as cursor:
            # 同時關聯統計每個客戶名下的報價單/訂單數量
            cursor.execute("""
                SELECT c.id, c.company_name, c.short_name, c.tax_code, c.contact_person, 
                       c.phone, c.email, c.address, c.default_tier, c.payment_terms, c.notes,
                       (SELECT COUNT(*) FROM tb_quotation q WHERE q.customer_id = c.id OR q.customer_name = c.company_name) as quote_count
                FROM tb_customer c
                ORDER BY c.id DESC LIMIT 200
            """)
            rows = cursor.fetchall()
            conn.close()
            return rows
    except Exception as e:
        print("獲取客戶列表失敗:", e)
        return []

@app.get("/api/customers/{customer_id}/quotes")
def get_customer_quotes(customer_id: int):
    """【核心突破】：穿透查詢指定客戶名下的所有報價單/訂單歷史"""
    try:
        conn = get_db()
        with conn.cursor() as cursor:
            cursor.execute("SELECT company_name, short_name FROM tb_customer WHERE id = %s", (customer_id,))
            cust = cursor.fetchone()
            if not cust:
                conn.close()
                return []
            
            c_name = cust['company_name']
            cursor.execute("""
                SELECT id, quote_no, customer_name, customer_tier, quote_currency, 
                       final_total_amount, status, created_at
                FROM tb_quotation
                WHERE customer_id = %s OR customer_name = %s
                ORDER BY id DESC
            """, (customer_id, c_name))
            rows = cursor.fetchall()
            for r in rows:
                if 'created_at' in r and r['created_at']:
                    r['created_at'] = str(r['created_at'])
            conn.close()
            return rows
    except Exception as e:
        print("獲取客戶訂單失敗:", e)
        return []

@app.post("/api/customers/save")
def save_customer(cust: CustomerModel):
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            if cust.id and cust.id > 0:
                cursor.execute("""
                    UPDATE tb_customer SET
                        company_name = %s, short_name = %s, tax_code = %s,
                        contact_person = %s, phone = %s, email = %s,
                        address = %s, default_tier = %s, payment_terms = %s, notes = %s
                    WHERE id = %s
                """, (
                    cust.company_name, cust.short_name, cust.tax_code,
                    cust.contact_person, cust.phone, cust.email,
                    cust.address, cust.default_tier, cust.payment_terms, cust.notes, cust.id
                ))
                cid = cust.id
            else:
                cursor.execute("""
                    INSERT INTO tb_customer (
                        company_name, short_name, tax_code, contact_person,
                        phone, email, address, default_tier, payment_terms, notes
                    ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """, (
                    cust.company_name, cust.short_name, cust.tax_code,
                    cust.contact_person, cust.phone, cust.email,
                    cust.address, cust.default_tier, cust.payment_terms, cust.notes
                ))
                cid = cursor.lastrowid
            conn.close()
            return {"status": "success", "id": cid, "message": "客戶檔案儲存成功！"}
    except Exception as e:
        if conn: conn.close()
        raise HTTPException(status_code=500, detail=str(e))

@app.delete("/api/customers/{customer_id}")
def delete_customer(customer_id: int):
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            cursor.execute("DELETE FROM tb_customer WHERE id = %s", (customer_id,))
            conn.close()
            return {"status": "success", "message": "客戶刪除成功！"}
    except Exception as e:
        if conn: conn.close()
        raise HTTPException(status_code=500, detail=str(e))

# ================= 7. 產品庫、匯率與歷史單據 API =================
@app.get("/api/products")
def get_products():
    try:
        conn = get_db()
        with conn.cursor() as cursor:
            cursor.execute("""
                SELECT p.id, p.sku, p.brand, p.category, p.model, p.spec, p.power_w, 
                       p.pcs_per_pallet, p.pcs_per_container, p.warranty_years,
                       COALESCE(
                           (SELECT unit_price FROM tb_price_matrix WHERE product_id = p.id AND customer_tier = '中型經銷商' LIMIT 1),
                           (SELECT unit_price FROM tb_price_matrix WHERE product_id = p.id LIMIT 1),
                           0
                       ) as default_base_usd
                FROM tb_product p
                WHERE p.is_active = 1
                ORDER BY p.category DESC, p.brand ASC, p.model ASC
            """)
            rows = cursor.fetchall()
            conn.close()
            return rows
    except Exception as e:
        print("獲取產品失敗:", e)
        return []

@app.get("/api/rates")
def get_rates():
    rates = {
        "USD_VND": 25967.0,
        "CNY_VND": 3874.0,
        "USD_CNY": 6.70,
        "USD_THB": 32.05,
        "USD_MYR": 3.9855
    }
    try:
        conn = get_db()
        with conn.cursor() as cursor:
            cursor.execute("SELECT pair, rate FROM tb_exchange_rate")
            rows = cursor.fetchall()
            for r in rows:
                rates[r['pair']] = float(r['rate'])
            conn.close()
    except Exception as e:
        print("獲取匯率失敗:", e)
    return rates

@app.get("/api/quotes/list")
def list_quotes():
    try:
        conn = get_db()
        with conn.cursor() as cursor:
            cursor.execute("""
                SELECT id, quote_no, customer_name, customer_tier, quote_currency, 
                       final_total_amount, status, created_at
                FROM tb_quotation
                ORDER BY id DESC LIMIT 500
            """)
            rows = cursor.fetchall()
            for r in rows:
                if 'created_at' in r and r['created_at']:
                    r['created_at'] = str(r['created_at'])
            conn.close()
            return rows
    except Exception as e:
        print("獲取報價列表失敗:", e)
        return []

@app.get("/api/quotes/{quote_id}")
def get_quote_detail(quote_id: int):
    try:
        conn = get_db()
        with conn.cursor() as cursor:
            cursor.execute("SELECT * FROM tb_quotation WHERE id = %s", (quote_id,))
            header = cursor.fetchone()
            if not header:
                conn.close()
                raise HTTPException(status_code=404, detail="報價單不存在")
            if 'created_at' in header and header['created_at']:
                header['created_at'] = str(header['created_at'])

            cursor.execute("""
                SELECT id, product_id, item_name, item_model, item_spec, unit, quantity, 
                       applied_tier, base_price, custom_markup, quote_unit_price, subtotal_amount, note
                FROM tb_quotation_item
                WHERE quotation_id = %s
            """, (quote_id,))
            items = cursor.fetchall()
            conn.close()
            return {"header": header, "items": items}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# ================= 8. 報價單儲存模型 (另存為新單與防覆蓋) =================
class EditableItem(BaseModel):
    product_id: Optional[int] = 0
    name: str
    spec: str = ""
    unit: str = "台 / Bộ"
    quantity: int = 1
    base_price: float = 0.0
    custom_markup: Optional[float] = None
    unit_price: float = 0.0
    subtotal: float = 0.0
    note: str = ""

class QuoteSaveUpdateRequest(BaseModel):
    quote_id: Optional[int] = None
    save_as_new: Optional[bool] = False
    layout_template_id: Optional[int] = 1
    customer_id: Optional[int] = 0
    quote_no: str
    customer_name: str
    contact_person: str = ""
    phone_email: str = ""
    sales_rep: str = "Thanh Bình 清平"
    customer_tier: str = "中型經銷商"
    delivery_scenario: str = "越南本地倉"
    trade_term: str = "DAP"
    target_currency: str = "VND"
    vat_rate_pct: float = 10.0
    valid_days: int = 3
    markup_pct: float = 0.0
    terms_text: str = ""
    terms_list: Optional[List[str]] = []
    lang_mode: str = "vi_zh"
    template_style: str = "standard"
    custom_title: Optional[str] = "BẢNG BÁO GIÁ 商業報價單"
    custom_greeting: Optional[str] = "Chúng tôi xin gửi đến Quý khách hàng bảng báo giá như sau / 我司現向貴司呈報以下報價表:"
    
    company_logo_data: Optional[str] = None
    company_name_vi: Optional[str] = "CÔNG TY TNHH TẬP ĐOÀN CÔNG NGHỆ GALAXY VIỆT NAM"
    company_name_cn: Optional[str] = "銀河科技集團（越南）有限公司"
    company_tax_code: Optional[str] = "2301296587"
    company_contact_info: Optional[str] = "Điện thoại: 0325.609.218 | Email: galaxyteck6886@gmail.com"
    company_address: Optional[str] = "Số 36 Yết Kiêu, Phường Kinh Bắc, Tỉnh Bắc Ninh, Việt Nam"
    company_bank_info: Optional[str] = "STK (Techcombank): 314431 - Ngân hàng TMCP Kỹ thương Việt Nam"
    
    status: str = "已生效"
    total_before_tax: float = 0.0
    vat_amount: float = 0.0
    final_total_after_tax: float = 0.0
    items: List[EditableItem]

@app.post("/api/quote/save_or_update")
def save_or_update_quote(req: QuoteSaveUpdateRequest):
    if not req.items:
        raise HTTPException(status_code=400, detail="明細列表不可為空")

    try:
        conn = get_db()
        with conn.cursor() as cursor:
            total_before_tax = Decimal(str(req.total_before_tax))
            final_total = Decimal(str(req.final_total_after_tax))
            terms_json_str = json.dumps(req.terms_list, ensure_ascii=False)

            cursor.execute("SELECT pair, rate FROM tb_exchange_rate")
            rate_map = {r['pair']: Decimal(str(r['rate'])) for r in cursor.fetchall()}
            rate_usd_vnd = rate_map.get("USD_VND", Decimal("25967.0"))
            rate_usd_cny = rate_map.get("USD_CNY", Decimal("6.70"))

            # 自定義物料自動入庫 (UUID8 確保絕對唯一)
            for it in req.items:
                clean_name = it.name.strip()
                if (not it.product_id or it.product_id == 0) and clean_name and clean_name != '自定義產品品名 / Description':
                    clean_spec = it.spec.strip() if it.spec else ''
                    cursor.execute("SELECT id FROM tb_product WHERE model = %s AND spec = %s LIMIT 1", (clean_name, clean_spec))
                    exist_prod = cursor.fetchone()

                    if exist_prod:
                        it.product_id = exist_prod['id']
                    else:
                        safe_tag = re.sub(r'[^\w\s-]', '', clean_name).replace(' ', '_')[:18]
                        new_sku = f"CUST-{safe_tag}-{uuid.uuid4().hex[:8].upper()}"
                        cursor.execute("""
                            INSERT INTO tb_product (
                                sku, brand, category, model, spec,
                                pcs_per_pallet, pcs_per_container, warranty_years, is_active
                            ) VALUES (%s, '通用自定義', '通用工程輔材', %s, %s, 1, 1, 1, 1)
                        """, (new_sku, clean_name, clean_spec))
                        it.product_id = cursor.lastrowid

                        base_dec = Decimal(str(it.base_price)) if it.base_price > 0 else Decimal("0.00")
                        if req.target_currency == "VND" and rate_usd_vnd > 0:
                            base_usd = (base_dec / rate_usd_vnd).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                        elif req.target_currency == "CNY" and rate_usd_cny > 0:
                            base_usd = (base_dec / rate_usd_cny).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                        else:
                            base_usd = base_dec.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)

                        for reg in ['越南本地倉', '中國出港']:
                            for tier in ['大型經銷商', '中型經銷商', '中型安裝商', '小型安裝商']:
                                cursor.execute("""
                                    INSERT INTO tb_price_matrix (
                                        product_id, region, trade_term, customer_tier, currency, unit_price
                                    ) VALUES (%s, %s, 'DAP', %s, 'USD', %s)
                                    ON DUPLICATE KEY UPDATE unit_price=VALUES(unit_price)
                                """, (it.product_id, reg, tier, float(base_usd)))

            final_quote_no = req.quote_no.strip()
            
            # 若不是強制另存為新單，且具備明確的 quote_id，則更新原單據
            if not req.save_as_new and req.quote_id and req.quote_id > 0:
                quote_id = req.quote_id
                cursor.execute("""
                    UPDATE tb_quotation SET
                        quote_no = %s, layout_template_id = %s, customer_id = %s, customer_name = %s, contact_person = %s,
                        phone_email = %s, sales_rep = %s, customer_tier = %s,
                        destination = %s, quote_currency = %s, markup_rate = %s,
                        tax_rate = %s, final_total_amount = %s, terms_text = %s, terms_json = %s,
                        valid_days = %s, lang_mode = %s, template_style = %s,
                        custom_title = %s, custom_greeting = %s,
                        company_logo_data = %s, company_name_vi = %s, company_name_cn = %s,
                        company_tax_code = %s, company_contact_info = %s,
                        company_address = %s, company_bank_info = %s, status = %s
                    WHERE id = %s
                """, (
                    final_quote_no, req.layout_template_id, req.customer_id, req.customer_name, req.contact_person,
                    req.phone_email, req.sales_rep, req.customer_tier,
                    f"{req.delivery_scenario} ({req.trade_term})", req.target_currency,
                    req.markup_pct, req.vat_rate_pct, float(final_total),
                    req.terms_text, terms_json_str, req.valid_days, req.lang_mode, req.template_style,
                    req.custom_title, req.custom_greeting,
                    req.company_logo_data, req.company_name_vi, req.company_name_cn,
                    req.company_tax_code, req.company_contact_info,
                    req.company_address, req.company_bank_info, req.status, quote_id
                ))
                cursor.execute("DELETE FROM tb_quotation_item WHERE quotation_id = %s", (quote_id,))
                msg = f"原報價單 [{final_quote_no}] 已成功更新儲存！"
            else:
                # 建立全新的獨立單據 (INSERT)
                cursor.execute("SELECT id FROM tb_quotation WHERE quote_no = %s LIMIT 1", (final_quote_no,))
                if cursor.fetchone():
                    now_date = datetime.datetime.now().strftime('%Y.%m.%d')
                    prefix = f"BG-{now_date}"
                    cursor.execute("SELECT quote_no FROM tb_quotation WHERE quote_no LIKE %s", (f"{prefix}%",))
                    all_existing = cursor.fetchall()
                    max_seq = 0
                    for row_it in all_existing:
                        m = re.search(r'-(\d+)$', row_it.get('quote_no', ''))
                        if m:
                            s = int(m.group(1))
                            if s > max_seq: max_seq = s
                    final_quote_no = f"{prefix}-{str(max_seq + 1).zfill(3)}"

                cursor.execute("""
                    INSERT INTO tb_quotation (
                        quote_no, layout_template_id, customer_id, customer_name, contact_person, phone_email, sales_rep,
                        customer_tier, destination, quote_currency, exchange_rate,
                        total_base_cost_usd, markup_rate, tax_rate, final_total_amount,
                        terms_text, terms_json, valid_days, lang_mode, template_style,
                        custom_title, custom_greeting,
                        company_logo_data, company_name_vi, company_name_cn,
                        company_tax_code, company_contact_info, company_address, company_bank_info, status
                    ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """, (
                    final_quote_no, req.layout_template_id, req.customer_id, req.customer_name, req.contact_person, req.phone_email, req.sales_rep,
                    req.customer_tier, f"{req.delivery_scenario} ({req.trade_term})", req.target_currency,
                    1.0, 0.0, req.markup_pct, req.vat_rate_pct, float(final_total),
                    req.terms_text, terms_json_str, req.valid_days, req.lang_mode, req.template_style,
                    req.custom_title, req.custom_greeting,
                    req.company_logo_data, req.company_name_vi, req.company_name_cn,
                    req.company_tax_code, req.company_contact_info, req.company_address, req.company_bank_info, req.status
                ))
                quote_id = cursor.lastrowid
                msg = f"全新報價單已成功建檔保存！[單號: {final_quote_no}]"

            for it in req.items:
                cursor.execute("""
                    INSERT INTO tb_quotation_item (
                        quotation_id, product_id, item_name, item_model, item_spec,
                        unit, quantity, applied_tier, base_price, custom_markup, quote_unit_price,
                        subtotal_amount, note
                    ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """, (
                    quote_id, it.product_id or 0, it.name, it.name, it.spec,
                    it.unit, it.quantity, req.customer_tier, it.base_price, it.custom_markup, it.unit_price,
                    it.subtotal, it.note
                ))

            conn.close()
            return {
                "status": "success",
                "message": msg,
                "quote_id": quote_id,
                "quote_no": final_quote_no,
                "final_total_after_tax": float(final_total)
            }
    except Exception as e:
        if conn: conn.close()
        print("資料庫保存異常詳情:", str(e))
        raise HTTPException(status_code=500, detail=f"資料庫保存出錯: {str(e)}")

# ================= 9. 前端單頁應用 (CRM 訂單穿透與一鍵調出維護) =================
@app.get("/", response_class=HTMLResponse)
def index():
    return """
    <!DOCTYPE html>
    <html lang="zh-TW">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>GalaxyTeck 智慧快速報價系統 v3.5.0</title>
      
      <script src="https://cdn.tailwindcss.com"></script>
      <script src="https://unpkg.com/vue@3/dist/vue.global.js"></script>
      <script>
        if (typeof Vue === 'undefined') {
          document.write('<script src="https://cdn.jsdelivr.net/npm/vue@3/dist/vue.global.prod.js"><\\/script>');
        }
      </script>

      <style>
        [v-cloak] { display: none !important; }

        .top-logo-img { max-height: 32px !important; width: auto !important; object-fit: contain !important; }
        .sheet-logo-img { max-height: 48px !important; width: auto !important; object-fit: contain !important; }

        body {
          font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
          background-color: #f8fafd;
          color: #1f2937;
        }

        @page {
          size: A4 portrait;
          margin: 12mm 14mm 12mm 14mm;
        }

        @media print {
          .no-print, header, nav, .print-hidden { display: none !important; }
          body { background: white !important; font-size: 10pt !important; line-height: 1.3 !important; margin: 0 !important; padding: 0 !important; }
          .print-sheet { width: 100% !important; max-width: 100% !important; margin: 0 !important; padding: 0 !important; border: none !important; box-shadow: none !important; background: white !important; }
          input, textarea, select { border: none !important; background: transparent !important; resize: none !important; padding: 0 !important; box-shadow: none !important; }
          table { page-break-inside: auto !important; width: 100% !important; border-collapse: collapse !important; }
          thead { display: table-header-group !important; }
          tr { page-break-inside: avoid !important; page-break-after: auto !important; }
          .keep-together { page-break-inside: avoid !important; }
          .compact-mode .header-box { margin-bottom: 2mm !important; padding-bottom: 1.5mm !important; }
          .compact-mode .meta-grid { margin-bottom: 2mm !important; }
          .compact-mode .item-table th, .compact-mode .item-table td { padding-top: 2px !important; padding-bottom: 2px !important; }
          .compact-mode .terms-box { margin-bottom: 2.5mm !important; padding: 2mm !important; }
          .compact-mode .signature-box { padding-top: 2mm !important; }
          .compact-mode .sig-space { height: 14mm !important; }
        }

        .a4-preview-card {
          width: 210mm;
          min-height: 297mm;
          margin: 0 auto;
          background: white;
          box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -2px rgba(0, 0, 0, 0.05), 0 10px 25px -5px rgba(26, 115, 232, 0.05);
          border: 1px solid #e2e8f0;
          border-radius: 12px;
          padding: 12mm 14mm;
          box-sizing: border-box;
          position: relative;
        }

        .auto-expand-text { field-sizing: content; resize: none; }
      </style>
    </head>
    <body class="bg-[#f8fafd] text-[#1f2937] min-h-screen">
      <div id="app" v-cloak class="py-4 px-2 md:px-6 max-w-[1360px] mx-auto">
        
        <!-- Google 企業級頂部主導覽列 -->
        <header class="no-print bg-white rounded-2xl shadow-sm border border-[#e0e3e7] p-3.5 mb-5 flex flex-col xl:flex-row justify-between items-center gap-3">
          
          <div class="flex items-center gap-3.5">
            <div class="relative group cursor-pointer" title="點擊更換公司 Logo (自動永久儲存至資料庫)">
              <label class="cursor-pointer block">
                <div v-if="companyInfo.logo_data" class="h-10 px-2.5 py-1 border border-slate-200 rounded-xl flex items-center bg-white hover:bg-slate-50 transition shadow-sm">
                  <img :src="companyInfo.logo_data" class="top-logo-img">
                </div>
                <div v-else class="h-10 px-3.5 bg-red-600 text-white rounded-xl font-bold text-sm tracking-wide shadow-sm hover:bg-red-700 transition flex items-center gap-2">
                  <span>GalaxyTECK</span>
                  <span class="text-xs bg-white/20 px-1.5 py-0.5 rounded">更換</span>
                </div>
                <input type="file" @change="handleLogoUploadAndPersist" accept="image/*" class="hidden">
              </label>
            </div>

            <div>
              <div class="flex items-center gap-2">
                <span class="text-base font-bold text-slate-800 tracking-tight">智慧快速報價工作台</span>
                <span class="px-2 py-0.5 bg-blue-50 text-blue-600 border border-blue-200 text-[10px] font-mono font-bold rounded-full">v3.5.0 CRM Pro</span>
              </div>
              <div class="text-[11px] text-slate-500 flex items-center gap-2 mt-0.5">
                <span>銀河科技集團(越南)有限公司</span>
                <span class="text-slate-300">|</span>
                <span>MST: 2301296587</span>
              </div>
            </div>
          </div>

          <div class="flex flex-wrap items-center gap-2.5 bg-slate-100 p-1.5 rounded-xl border border-slate-200">
            <div class="flex items-center bg-white rounded-lg p-0.5 shadow-sm border border-slate-200 text-xs">
              <button type="button" @click="viewMode = 'customer'" :class="['px-3.5 py-1.5 rounded-md font-bold transition flex items-center gap-1.5', viewMode === 'customer' ? 'bg-blue-600 text-white shadow-sm' : 'text-slate-600 hover:text-slate-900']">
                👔 客戶對外版
              </button>
              <button type="button" @click="viewMode = 'finance'" :class="['px-3.5 py-1.5 rounded-md font-bold transition flex items-center gap-1.5', viewMode === 'finance' ? 'bg-amber-600 text-white shadow-sm' : 'text-slate-600 hover:text-slate-900']">
                📊 內部財務版
              </button>
            </div>

            <div class="flex items-center bg-white rounded-lg p-0.5 shadow-sm border border-slate-200 text-xs">
              <span class="text-slate-400 font-medium px-2 text-[11px]">A4版面:</span>
              <button type="button" @click="printDensity = 'auto'" :class="['px-2.5 py-1 rounded font-bold transition text-[11px]', printDensity === 'auto' ? 'bg-slate-800 text-white' : 'text-slate-600']">
                智能自適應
              </button>
              <button type="button" @click="printDensity = 'ultra'" :class="['px-2.5 py-1 rounded font-bold transition text-[11px]', printDensity === 'ultra' ? 'bg-slate-800 text-white' : 'text-slate-600']">
                緊湊單頁
              </button>
            </div>
          </div>

          <div class="flex flex-wrap items-center gap-2">
            <button type="button" @click="showTemplateDrawer = !showTemplateDrawer" class="px-3 py-2 bg-white hover:bg-slate-50 text-blue-600 rounded-xl text-xs font-semibold border border-slate-200 hover:border-blue-500 transition shadow-sm flex items-center gap-1.5">
              🎨 報價版式與抬頭設置
            </button>
            <button type="button" @click="openCustomerModal" class="px-3 py-2 bg-white hover:bg-slate-50 text-emerald-600 rounded-xl text-xs font-semibold border border-slate-200 hover:border-emerald-500 transition shadow-sm flex items-center gap-1.5">
              👥 CRM 客戶管理中心
            </button>
            <button type="button" @click="openHistoryModal" class="px-3 py-2 bg-white hover:bg-slate-50 text-slate-700 rounded-xl text-xs font-semibold border border-slate-200 hover:border-slate-800 transition shadow-sm flex items-center gap-1.5">
              📂 報價歷史庫 ({{ historyList.length }})
            </button>
          </div>
        </header>

        <!-- 財務審核版內部機密提示條 -->
        <div v-if="viewMode === 'finance'" class="no-print mb-4 p-3 bg-amber-50 border border-amber-300 rounded-xl text-amber-800 flex justify-between items-center text-xs">
          <div class="flex items-center gap-2.5 font-bold">
            <span class="px-2 py-0.5 bg-amber-600 text-white rounded-md text-[10px]">內部機密</span>
            <span>當前為【內部財務審核版】：顯示基準採購成本、全局與單品調價率、行毛利額與綜合利潤率。嚴禁發送給客戶！</span>
          </div>
          <button @click="viewMode = 'customer'" class="text-blue-600 hover:underline font-semibold">切換回客戶對外版</button>
        </div>

        <!-- 商業與調價參數卡片 -->
        <div class="no-print bg-white p-5 rounded-2xl shadow-sm border border-slate-200 mb-5 space-y-4">
          <div class="text-xs font-bold text-slate-800 uppercase tracking-wider border-b border-slate-100 pb-2.5 flex justify-between items-center">
            <span class="flex items-center gap-1.5">
              <span class="w-2 h-2 bg-blue-600 rounded-full"></span>
              商業計價與即時交付參數
            </span>
            <span class="text-blue-600 font-mono font-medium text-xs">
              匯率基準：1 USD = {{ rates['USD_VND'] || 25967 }} VND | 1 USD = {{ rates['USD_CNY'] || 6.70 }} CNY
            </span>
          </div>

          <div class="grid grid-cols-1 md:grid-cols-4 gap-3 text-xs">
            <div>
              <div class="flex justify-between items-center mb-1">
                <label class="font-medium text-slate-500">報價單號</label>
                <button type="button" @click="fetchNextQuoteNo" class="text-[10px] text-blue-600 hover:underline">換新單號</button>
              </div>
              <input v-model="form.quote_no" class="w-full border border-slate-200 rounded-xl px-3 py-2 font-mono font-bold text-blue-600 bg-slate-50 focus:bg-white focus:border-blue-500 outline-none transition">
            </div>
            
            <div class="relative">
              <div class="flex justify-between items-center mb-1">
                <label class="font-medium text-slate-500">客戶名稱</label>
                <button type="button" @click="openCustomerModal" class="text-[11px] text-blue-600 hover:underline font-bold">CRM庫選客戶</button>
              </div>
              <input v-model="form.customer_name" @focus="showCustomerDropdown = true" placeholder="輸入或點選客戶名稱..." class="w-full border border-slate-200 rounded-xl px-3 py-2 font-semibold text-slate-800 focus:border-blue-500 outline-none transition">
              
              <div v-if="showCustomerDropdown && matchedCustomers.length > 0" class="absolute left-0 right-0 top-full mt-1.5 bg-white border border-slate-200 rounded-xl shadow-lg z-40 max-h-52 overflow-y-auto p-1">
                <div v-for="c in matchedCustomers" :key="c.id" @click="selectCustomerToForm(c)" class="p-2.5 hover:bg-slate-50 rounded-lg cursor-pointer transition">
                  <div class="font-bold text-slate-800 text-xs">{{ c.company_name }}</div>
                  <div class="text-[11px] text-slate-500 mt-0.5">{{ c.contact_person }} · {{ c.phone }} · 階梯: {{ c.default_tier }}</div>
                </div>
              </div>
            </div>

            <div>
              <label class="block font-medium text-slate-500 mb-1">聯絡人</label>
              <input v-model="form.contact_person" class="w-full border border-slate-200 rounded-xl px-3 py-2 focus:border-blue-500 outline-none transition" placeholder="聯絡人姓名">
            </div>
            <div>
              <label class="block font-medium text-slate-500 mb-1">電話 / 郵箱</label>
              <input v-model="form.phone_email" class="w-full border border-slate-200 rounded-xl px-3 py-2 focus:border-blue-500 outline-none transition" placeholder="電話或電子郵箱">
            </div>
          </div>

          <div class="grid grid-cols-2 md:grid-cols-5 gap-3 text-xs">
            <div>
              <label class="block font-medium text-slate-500 mb-1">交付場景</label>
              <select v-model="form.delivery_scenario" class="w-full border border-slate-200 rounded-xl px-3 py-2 bg-white focus:border-blue-500 outline-none transition">
                <option value="越南本地倉">越南本地倉 (DAP)</option>
                <option value="泰國本地倉">泰國本地倉 (EXW)</option>
                <option value="中國出港">中國出港 (FOB)</option>
              </select>
            </div>
            <div>
              <label class="block font-medium text-slate-500 mb-1">價格階梯</label>
              <select v-model="form.customer_tier" class="w-full border border-slate-200 rounded-xl px-3 py-2 bg-white focus:border-blue-500 outline-none transition font-medium">
                <option value="大型經銷商">大型經銷商 (整櫃批量)</option>
                <option value="中型經銷商">中型經銷商 (>500萬)</option>
                <option value="中型安裝商">中型安裝商 (整托)</option>
                <option value="小型安裝商">小型安裝商 (散單)</option>
              </select>
            </div>
            <div>
              <label class="block font-medium text-slate-500 mb-1">結算幣種</label>
              <select v-model="form.target_currency" @change="onCurrencyChange" class="w-full border border-slate-200 rounded-xl px-3 py-2 bg-white font-bold text-blue-600 focus:border-blue-500 outline-none transition">
                <option value="VND">VND (越南盾 ₫)</option>
                <option value="CNY">CNY (人民幣 ¥)</option>
                <option value="USD">USD ($ 美元)</option>
              </select>
            </div>
            <div>
              <label class="block font-medium text-slate-500 mb-1">增值稅 VAT (%)</label>
              <input type="number" step="1" v-model.number="form.vat_rate_pct" class="w-full border border-slate-200 rounded-xl px-3 py-2 text-right font-mono focus:border-blue-500 outline-none transition">
            </div>
            
            <div class="bg-amber-50 p-2 rounded-xl border border-amber-300">
              <label class="block font-bold text-amber-800 mb-0.5 flex justify-between text-[11px]">
                <span>⚡ 全局調價比例 (%)</span>
                <span class="text-amber-600">整體聯動</span>
              </label>
              <input type="number" step="0.5" v-model.number="form.markup_pct" class="w-full border border-amber-300 rounded-lg px-2.5 py-1 text-right font-mono font-bold text-sm text-amber-800 bg-white outline-none" placeholder="0">
            </div>
          </div>

          <div class="flex flex-wrap gap-2.5 pt-3 border-t border-slate-100 items-center">
            <button type="button" @click="openSearchModal" class="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-xl text-xs font-bold transition flex items-center gap-1.5 shadow-sm">
              🔍 檢索產品庫 (快捷選型)
            </button>
            <div class="relative flex-1 min-w-[280px]">
              <input type="text" v-model="quickSearchKeyword" @keyup.enter="quickAddFirstMatch" placeholder="直接輸入產品型號、品牌或規格關鍵字 (按 Enter 鍵可快速選入首項)..." class="w-full border border-slate-200 rounded-full px-4 py-2 text-xs focus:ring-2 focus:ring-blue-100 focus:border-blue-500 outline-none transition shadow-sm">
              <span v-if="filteredProductList.length > 0 && quickSearchKeyword" class="absolute right-3 top-2 text-[10px] bg-blue-50 text-blue-600 font-bold px-2 py-0.5 rounded-full">
                匹配 {{ filteredProductList.length }} 項
              </span>
            </div>
            <button type="button" @click="resetAllCustomMarkups" v-if="hasAnyCustomMarkup" class="px-3 py-2 bg-amber-50 hover:bg-amber-100 text-amber-800 rounded-xl text-xs font-bold transition border border-amber-300">
              ↺ 恢復全局調價
            </button>
            <button type="button" @click="addCustomBlankRow" class="px-4 py-2 bg-slate-800 hover:bg-slate-900 text-white rounded-xl text-xs font-bold transition shadow-sm">
              + 插入自定義物料行
            </button>
          </div>
        </div>

        <!-- 報價版式與抬頭自定義管理抽屜 (列印隱藏) -->
        <div v-show="showTemplateDrawer" style="display: none;" class="no-print bg-white border border-blue-200 p-6 rounded-2xl shadow-md mb-5 space-y-4">
          <div class="flex justify-between items-center border-b border-slate-100 pb-3">
            <h3 class="text-sm font-bold text-blue-600 tracking-wide flex items-center gap-2">
              <span>🎨 報價版式範本庫與公司抬頭完全自定義設置 (Layout & Branding)</span>
            </h3>
            <button type="button" @click="showTemplateDrawer = false" class="text-sm text-slate-500 hover:text-slate-800 font-bold">收起 ✕</button>
          </div>

          <div class="p-3.5 bg-slate-50 rounded-xl border border-slate-200 flex flex-col md:flex-row justify-between items-center gap-3">
            <div class="flex items-center gap-2.5 flex-1 w-full">
              <span class="text-xs font-bold text-slate-800 whitespace-nowrap">套用版式範本:</span>
              <select v-model="selectedLayoutTemplateId" @change="applyLayoutTemplateById($event.target.value)" class="border border-slate-200 rounded-xl px-3 py-2 bg-white text-blue-600 font-bold text-xs flex-1 max-w-md outline-none focus:border-blue-500">
                <option v-for="tpl in layoutTemplatesList" :key="tpl.id" :value="tpl.id">
                  {{ tpl.template_name }} {{ tpl.is_default ? '(預設)' : '' }}
                </option>
              </select>
            </div>
            <div class="flex items-center gap-2 w-full md:w-auto justify-end">
              <button type="button" @click="promptSaveNewLayoutTemplate" class="px-3.5 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-xl text-xs font-bold shadow-sm">
                + 另存為新版式
              </button>
              <button type="button" @click="updateCurrentLayoutTemplate" class="px-3.5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-bold shadow-sm">
                覆蓋儲存當前版式
              </button>
              <button v-if="layoutTemplatesList.length > 1 && selectedLayoutTemplateId" type="button" @click="deleteCurrentLayoutTemplate" class="px-3 py-2 bg-rose-50 hover:bg-rose-100 text-rose-600 rounded-xl text-xs font-bold border border-rose-200">
                刪除此版式
              </button>
            </div>
          </div>

          <div class="grid grid-cols-1 md:grid-cols-2 gap-3.5 text-xs">
            <div>
              <label class="block font-bold text-slate-800 mb-1">公司全稱 (越文/英文法定主抬頭 - 自然換行)</label>
              <textarea v-model="companyInfo.name_vi" rows="2" class="w-full border border-slate-200 rounded-xl px-3 py-2 bg-white font-bold leading-snug focus:border-blue-500 outline-none"></textarea>
            </div>
            <div>
              <label class="block font-bold text-slate-800 mb-1">公司全稱 (中文法定全稱/業務附抬頭)</label>
              <textarea v-model="companyInfo.name_cn" rows="2" class="w-full border border-slate-200 rounded-xl px-3 py-2 bg-white leading-snug focus:border-blue-500 outline-none"></textarea>
            </div>
          </div>

          <div class="grid grid-cols-1 md:grid-cols-2 gap-3 text-xs pt-3 border-t border-slate-100">
            <div>
              <label class="block font-bold text-slate-500 mb-1">企業代碼 / 稅號 (Mã số thuế / Tax code)</label>
              <input v-model="companyInfo.tax_code" class="w-full border border-slate-200 rounded-xl px-3 py-2 bg-white font-mono focus:border-blue-500 outline-none">
            </div>
            <div>
              <label class="block font-bold text-slate-500 mb-1">電話與電子郵件 (Điện thoại / Email)</label>
              <input v-model="companyInfo.contact_info" class="w-full border border-slate-200 rounded-xl px-3 py-2 bg-white focus:border-blue-500 outline-none">
            </div>
            <div>
              <label class="block font-bold text-slate-500 mb-1">法定登記地址 (Địa chỉ doanh nghiệp)</label>
              <input v-model="companyInfo.address" class="w-full border border-slate-200 rounded-xl px-3 py-2 bg-white focus:border-blue-500 outline-none">
            </div>
            <div>
              <label class="block font-bold text-slate-500 mb-1">銀行帳戶資訊 (Tài khoản ngân hàng / Bank info)</label>
              <input v-model="companyInfo.bank_info" class="w-full border border-slate-200 rounded-xl px-3 py-2 bg-white font-mono focus:border-blue-500 outline-none">
            </div>
          </div>

          <div class="grid grid-cols-1 md:grid-cols-3 gap-3 text-xs pt-3 border-t border-slate-100">
            <div>
              <label class="block font-bold text-slate-500 mb-1">單據主標題文字 (Title)</label>
              <input v-model="layoutConfig.title_text" class="w-full border border-slate-200 rounded-xl px-3 py-2 bg-white font-bold text-slate-800 focus:border-blue-500 outline-none">
            </div>
            <div class="md:col-span-2">
              <label class="block font-bold text-slate-500 mb-1">致辭敬語 (Greeting Text)</label>
              <input v-model="layoutConfig.greeting_text" class="w-full border border-slate-200 rounded-xl px-3 py-2 bg-white text-slate-700 focus:border-blue-500 outline-none">
            </div>
          </div>
        </div>

        <!-- 核心生命週期動作條 -->
        <div class="no-print max-w-[210mm] mx-auto mb-3 flex flex-col sm:flex-row justify-between items-center gap-2.5 px-1">
          <div class="flex items-center gap-2 text-xs">
            <span class="inline-flex items-center gap-1.5 px-2.5 py-1 bg-white border border-slate-200 rounded-lg font-mono font-bold text-slate-800 shadow-sm">
              <span class="w-2 h-2 rounded-full" :class="currentQuoteId ? 'bg-amber-500 animate-pulse' : 'bg-emerald-500'"></span>
              <span>單號: {{ form.quote_no || '未生成' }}</span>
            </span>
            <span v-if="currentQuoteId" class="px-2 py-0.5 bg-amber-50 border border-amber-200 text-amber-800 text-[11px] font-semibold rounded-md">
              (正在編輯已存檔單據 ID: {{ currentQuoteId }})
            </span>
            <span v-else class="px-2 py-0.5 bg-emerald-50 border border-emerald-200 text-emerald-700 text-[11px] font-semibold rounded-md">
              (全新獨立單據就緒)
            </span>
          </div>

          <div class="flex items-center gap-2">
            <button type="button" @click="handleSafeCreateNewQuote" class="px-3.5 py-2 bg-gradient-to-r from-teal-600 to-emerald-600 hover:from-teal-700 hover:to-emerald-700 text-white rounded-xl text-xs font-bold transition shadow-sm hover:shadow flex items-center gap-1.5" title="清空當前畫布，開啟下一張新流水號報價單">
              <span>📄✨</span>
              <span>新建空白報價單</span>
            </button>
            <button v-if="currentQuoteId" type="button" @click="saveToServer(true)" class="px-3.5 py-2 bg-sky-600 hover:bg-sky-700 text-white rounded-xl text-xs font-bold shadow-sm transition flex items-center gap-1.5" title="不覆蓋原歷史單據，自動領取新流水號另存為獨立新單據">
              <span>📋</span>
              <span>另存為新報價單</span>
            </button>
            <button type="button" @click="saveToServer(false)" class="px-3.5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-bold shadow-sm transition flex items-center gap-1.5">
              <span>💾</span>
              <span>{{ currentQuoteId ? '覆蓋更新當前單' : '報價單保存入庫' }}</span>
            </button>
            <button type="button" @click="printOffer" class="px-3.5 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-xl text-xs font-bold shadow-sm transition flex items-center gap-1.5">
              <span>🖨️</span>
              <span>列印 / 匯出</span>
            </button>
          </div>
        </div>

        <!-- 1:1 所見即所得 A4 報價單主體容器 -->
        <div :class="['a4-preview-card print-sheet text-slate-800 font-sans', getEffectiveDensityClass()]">
          
          <!-- 公司抬頭區 -->
          <div class="header-box border-b border-slate-300 pb-3 mb-3">
            <div class="flex justify-between items-start gap-4">
              <div class="flex-shrink-0 max-w-[240px]">
                <div v-if="companyInfo.logo_data" class="flex items-center">
                  <img :src="companyInfo.logo_data" alt="Company Logo" class="sheet-logo-img">
                </div>
                <div v-else>
                  <div class="flex items-center gap-1.5">
                    <span class="text-2xl font-black tracking-tight text-red-600">G</span>
                    <span class="text-xl font-black tracking-wider text-slate-900">GalaxyTECK</span>
                  </div>
                  <div class="text-[9px] tracking-widest text-slate-400 font-bold uppercase mt-0.5">SOLUTIONS WE OFFER</div>
                </div>
              </div>

              <div class="flex-1 min-w-0 text-right">
                <textarea v-model="companyInfo.name_vi" rows="1" class="auto-expand-text text-sm font-bold text-slate-900 text-right w-full bg-transparent border-b border-transparent focus:border-blue-400 outline-none leading-snug break-words" placeholder="公司法定全稱 (越文/英文)"></textarea>
                <textarea v-model="companyInfo.name_cn" rows="1" class="auto-expand-text text-[11px] font-medium text-slate-600 text-right w-full bg-transparent border-b border-transparent focus:border-blue-400 outline-none leading-tight break-words mt-0.5" placeholder="公司中文名稱"></textarea>
              </div>
            </div>

            <div class="mt-2 pt-2 border-t border-slate-100 grid grid-cols-2 text-[10px] text-slate-600 gap-x-4 gap-y-0.5">
              <div>
                <span class="font-semibold text-slate-700">Mã số thuế (Tax code):</span> 
                <input v-model="companyInfo.tax_code" class="bg-transparent border-b border-transparent focus:border-blue-400 outline-none font-mono">
              </div>
              <div class="text-right">
                <input v-model="companyInfo.contact_info" class="bg-transparent border-b border-transparent focus:border-blue-400 outline-none text-right w-full">
              </div>
              <div class="col-span-2">
                <span class="font-semibold text-slate-700">Địa chỉ:</span> 
                <input v-model="companyInfo.address" class="bg-transparent border-b border-transparent focus:border-blue-400 outline-none w-[85%]">
              </div>
              <div class="col-span-2">
                <input v-model="companyInfo.bank_info" class="bg-transparent border-b border-transparent focus:border-blue-400 outline-none w-full font-mono text-[9.5px]">
              </div>
            </div>
          </div>

          <!-- 主標題 -->
          <div class="text-center my-3">
            <div v-if="viewMode === 'finance'" class="inline-block px-2.5 py-0.5 bg-amber-100 border border-amber-300 text-amber-800 rounded font-bold text-[10px] uppercase mb-1">
              內部審批與財務核算專用 · 嚴禁對外交付
            </div>
            <input v-model="layoutConfig.title_text" class="text-lg font-black tracking-wider text-slate-900 text-center w-full bg-transparent border-b border-transparent focus:border-blue-400 outline-none">
          </div>

          <!-- 客戶信息與報價信息欄 -->
          <div class="meta-grid grid grid-cols-2 text-[11px] mb-3 divide-x border border-slate-300 divide-slate-300">
            <div class="p-2 space-y-1">
              <div class="font-bold border-b border-slate-200 pb-1 mb-1 text-slate-700 text-[10px] uppercase">THÔNG TIN KHÁCH HÀNG 客戶信息</div>
              <div class="flex items-center"><span class="w-24 text-slate-500 font-medium">Khách hàng 客戶:</span><input v-model="form.customer_name" class="flex-1 font-semibold text-slate-900 border-b border-dashed border-slate-200 focus:outline-none"></div>
              <div class="flex items-center"><span class="w-24 text-slate-500 font-medium">Người liên hệ:</span><input v-model="form.contact_person" class="flex-1 border-b border-dashed border-slate-200 focus:outline-none"></div>
              <div class="flex items-center"><span class="w-24 text-slate-500 font-medium">Điện thoại/Email:</span><input v-model="form.phone_email" class="flex-1 border-b border-dashed border-slate-200 focus:outline-none"></div>
            </div>

            <div class="p-2 space-y-1">
              <div class="font-bold border-b border-slate-200 pb-1 mb-1 text-slate-700 text-[10px] uppercase">THÔNG TIN BÁO GIÁ 報價信息</div>
              <div class="flex justify-between items-center"><span class="text-slate-500 font-medium">Số báo giá 單號:</span><input v-model="form.quote_no" class="text-right font-bold text-slate-900 font-mono border-b border-dashed border-slate-200 focus:outline-none"></div>
              <div class="flex justify-between"><span class="text-slate-500 font-medium">Ngày báo giá 日期:</span><span v-text="todayDate"></span></div>
              <div class="flex justify-between items-center"><span class="text-slate-500 font-medium">Người báo giá 報價人:</span><input v-model="form.sales_rep" class="text-right border-b border-dashed border-slate-200 focus:outline-none"></div>
              <div class="flex justify-between"><span class="text-slate-500 font-medium">Hiệu lực 有效期:</span><span class="font-semibold text-red-600"><span v-text="validUntilDate"></span> (<span v-text="form.valid_days"></span> ngày)</span></div>
            </div>
          </div>

          <!-- 內部財務專用：綜合利潤指標看板 -->
          <div v-if="viewMode === 'finance'" class="mb-3 p-2.5 bg-slate-50 border border-slate-300 rounded-lg keep-together">
            <div class="font-bold text-slate-800 text-[11px] mb-2 flex items-center justify-between border-b pb-1">
              <span>📊 內部利潤與成本指標分析 (Margin & Profit Summary)</span>
              <span class="text-blue-600 font-mono font-bold text-[10px]">全局調價比例: {{ form.markup_pct }}%</span>
            </div>
            <div class="grid grid-cols-2 md:grid-cols-4 gap-2 text-center font-mono">
              <div class="p-1.5 bg-white border rounded">
                <div class="text-[10px] text-slate-500 font-sans">總採購/基準成本</div>
                <div class="text-xs font-bold text-slate-800 mt-0.5">{{ formatCurrency(totalBaseCost) }}</div>
              </div>
              <div class="p-1.5 bg-white border rounded">
                <div class="text-[10px] text-slate-500 font-sans">對外稅前總營收</div>
                <div class="text-xs font-bold text-blue-700 mt-0.5">{{ formatCurrency(totalBeforeTax) }}</div>
              </div>
              <div class="p-1.5 bg-white border rounded">
                <div class="text-[10px] text-slate-500 font-sans">預計毛利總額 (Profit)</div>
                <div :class="['text-xs font-black mt-0.5', totalGrossProfit >= 0 ? 'text-emerald-600' : 'text-red-600']">
                  {{ formatCurrency(totalGrossProfit) }}
                </div>
              </div>
              <div class="p-1.5 bg-white border rounded">
                <div class="text-[10px] text-slate-500 font-sans">綜合毛利率 (Margin %)</div>
                <div :class="['text-xs font-black mt-0.5', grossMarginPercent >= 0 ? 'text-emerald-600' : 'text-red-600']">
                  {{ grossMarginPercent.toFixed(2) }}%
                </div>
              </div>
            </div>
          </div>

          <div class="text-[10px] text-slate-600 italic mb-1.5">
            <input v-model="layoutConfig.greeting_text" class="w-full bg-transparent border-b border-transparent focus:border-blue-400 outline-none italic">
          </div>

          <!-- 報價清單表格 -->
          <div class="overflow-x-auto mb-3">
            <table class="item-table w-full border-collapse text-[11px] border border-slate-300">
              <thead>
                <tr class="bg-slate-100 text-slate-700 text-center font-bold">
                  <th class="border border-slate-300 py-1.5 px-1 w-8">STT<br>序號</th>
                  <th class="border border-slate-300 py-1.5 px-2 text-left">Tên sản phẩm<br>產品名稱</th>
                  <th class="border border-slate-300 py-1.5 px-2 text-left">Thông số kỹ thuật<br>規格型號</th>
                  <th class="border border-slate-300 py-1.5 px-1 w-12 text-center">Đơn vị<br>單位</th>
                  <th class="border border-slate-300 py-1.5 px-1 w-12 text-center">SL<br>數量</th>
                  
                  <th v-if="viewMode === 'finance'" class="border border-slate-300 py-1.5 px-1 text-right w-20 bg-slate-200/50">
                    底價成本<br><span class="text-[9px] text-slate-500 font-normal">Base Cost</span>
                  </th>

                  <th class="no-print border border-slate-300 py-1.5 px-1 text-center w-16 bg-amber-50">
                    獨立調價(%)<br><span class="text-[8px] text-slate-400 font-normal">空=跟隨全局</span>
                  </th>

                  <th class="border border-slate-300 py-1.5 px-2 text-right w-24 bg-amber-50/40">
                    Đơn giá<br>對外單價
                  </th>

                  <th class="border border-slate-300 py-1.5 px-2 text-right w-28">Thành tiền<br>總價</th>

                  <th v-if="viewMode === 'finance'" class="border border-slate-300 py-1.5 px-1 text-right w-20 bg-emerald-50 text-emerald-900">
                    毛利貢獻<br><span class="text-[9px] font-normal">Profit</span>
                  </th>

                  <th class="no-print border border-slate-300 py-1.5 px-1 w-8">操作</th>
                </tr>
              </thead>
              <tbody>
                <tr v-for="(row, idx) in tableRows" :key="idx" class="hover:bg-slate-50 transition">
                  <td class="border border-slate-300 py-1 px-1 text-center font-mono" v-text="idx + 1"></td>
                  <td class="border border-slate-300 py-1 px-2 font-semibold text-slate-900">
                    <input v-model="row.name" class="w-full bg-transparent outline-none">
                  </td>
                  <td class="border border-slate-300 py-1 px-2 text-slate-600 font-mono text-[10px] leading-tight">
                    <input v-model="row.spec" class="w-full bg-transparent outline-none">
                  </td>
                  <td class="border border-slate-300 py-1 px-1 text-center text-slate-600">
                    <input v-model="row.unit" class="w-full text-center bg-transparent outline-none">
                  </td>
                  <td class="border border-slate-300 py-1 px-1 text-center font-mono">
                    <input type="number" min="1" v-model.number="row.quantity" class="w-10 text-center font-mono text-[11px] border border-transparent hover:border-slate-300 rounded focus:border-blue-500">
                  </td>

                  <td v-if="viewMode === 'finance'" class="border border-slate-300 py-1 px-1 text-right font-mono text-slate-500 bg-slate-50">
                    <input type="number" step="0.01" v-model.number="row.base_price" class="w-16 text-right font-mono border-b border-dashed border-slate-300 focus:border-blue-500 outline-none">
                  </td>

                  <td class="no-print border border-slate-300 py-1.5 px-1 text-center bg-amber-50/30">
                    <input type="number" step="0.5" v-model.number="row.custom_markup" :placeholder="form.markup_pct" class="w-12 border border-amber-300 rounded px-0.5 text-center font-mono text-[11px] text-amber-900 font-bold focus:bg-amber-100 outline-none">
                    <span v-if="row.custom_markup !== null && row.custom_markup !== undefined && row.custom_markup !== ''" class="block text-[7px] text-amber-600 font-bold leading-none mt-0.5">📌單獨</span>
                  </td>

                  <td class="border border-slate-300 py-1 px-2 text-right font-mono font-bold text-slate-900 bg-amber-50/20">
                    <input type="number" step="0.01" :value="calcRowFinalUnitPrice(row)" @change="onManualUnitPriceChange(row, $event.target.value)" class="w-20 text-right font-mono font-bold text-slate-900 bg-transparent border-b border-transparent focus:border-blue-500 outline-none">
                  </td>

                  <td class="border border-slate-300 py-1 px-2 text-right font-mono font-bold text-slate-900">
                    <span v-text="formatCurrency(calcRowSubtotal(row))"></span>
                  </td>

                  <td v-if="viewMode === 'finance'" class="border border-slate-300 py-1 px-1 text-right font-mono font-bold text-emerald-700 bg-emerald-50/40">
                    <span v-text="formatCurrency(calcRowProfit(row))"></span>
                  </td>

                  <td class="no-print border border-slate-300 py-1.5 px-1 text-center">
                    <button type="button" @click="removeRow(idx)" class="text-red-500 hover:text-red-700 font-bold px-1 py-0.5 text-xs">✕</button>
                  </td>
                </tr>

                <tr v-if="tableRows.length === 0">
                  <td :colspan="viewMode === 'finance' ? 11 : 9" class="text-center py-6 text-slate-400">目前清單為空，請點選上方按鈕搜尋並新增產品</td>
                </tr>
              </tbody>
            </table>
          </div>

          <!-- 稅金與結算匯總 -->
          <div class="keep-together flex justify-end mb-3">
            <div class="w-72 border border-slate-300 text-[11px] divide-y divide-slate-300 font-mono">
              <div class="flex justify-between py-1 px-2.5 bg-slate-50">
                <span class="font-sans font-bold text-slate-700">TỔNG TRƯỚC THUẾ 稅前總額:</span>
                <span class="font-bold text-slate-900"><span v-text="formatCurrency(totalBeforeTax)"></span> <span v-text="form.target_currency"></span></span>
              </div>
              <div class="flex justify-between py-1 px-2.5">
                <span class="font-sans text-slate-600">THUẾ VAT 增值稅 (<span v-text="form.vat_rate_pct"></span>%):</span>
                <span><span v-text="formatCurrency(vatAmount)"></span> <span v-text="form.target_currency"></span></span>
              </div>
              <div class="flex justify-between py-1.5 px-2.5 font-bold text-xs bg-blue-50 text-blue-900">
                <span class="font-sans">TỔNG SAU THUẾ 含稅總額:</span>
                <span><span v-text="formatCurrency(finalTotalAfterTax)"></span> <span v-text="form.target_currency"></span></span>
              </div>
            </div>
          </div>

          <!-- 結構化無滑動公文條款清單 -->
          <div class="terms-box keep-together mb-4 text-[10px] text-slate-700 bg-slate-50 p-3 border border-slate-300 rounded">
            <div class="flex justify-between items-center mb-1.5 border-b border-slate-200 pb-1">
              <div class="font-bold text-slate-900 text-[10.5px]">ĐIỀU KIỆN VÀ GHI CHÚ / 商務條款及備註:</div>
              <div class="no-print flex items-center gap-1.5">
                <button type="button" @click="addTermRow" class="text-[10px] px-2.5 py-0.5 bg-blue-600 hover:bg-blue-700 text-white rounded font-bold transition">
                  + 新增一條
                </button>
              </div>
            </div>

            <div class="space-y-1">
              <div v-for="(term, tIdx) in termsList" :key="tIdx" class="flex items-start gap-1 group">
                <span class="font-bold text-slate-600 font-mono text-[10px] pt-0.5">{{ tIdx + 1 }})</span>
                <input v-model="termsList[tIdx]" class="flex-1 bg-transparent border-b border-transparent hover:border-slate-300 focus:border-blue-400 outline-none text-[10px] text-slate-800 leading-normal" placeholder="輸入條款內容...">
                
                <div class="no-print opacity-0 group-hover:opacity-100 flex items-center gap-1 transition-opacity">
                  <button type="button" @click="moveTermUp(tIdx)" :disabled="tIdx === 0" class="text-slate-400 hover:text-slate-600 disabled:opacity-30">▲</button>
                  <button type="button" @click="moveTermDown(tIdx)" :disabled="tIdx === termsList.length - 1" class="text-slate-400 hover:text-slate-600 disabled:opacity-30">▼</button>
                  <button type="button" @click="removeTermRow(tIdx)" class="text-rose-500 hover:text-rose-700 font-bold px-1">✕</button>
                </div>
              </div>
            </div>
          </div>

          <!-- 簽章區 -->
          <div class="signature-box keep-together grid grid-cols-2 text-center text-[11px] pt-1 border-t border-slate-200">
            <div>
              <div class="font-bold text-slate-900 mb-1 uppercase">
                <span v-if="viewMode === 'customer'">XÁC NHẬN CỦA KHÁCH HÀNG<br><span class="font-normal text-[10px]">客戶確認簽字蓋章</span></span>
                <span v-else>NGƯỜI LẬP BIỂU / 製單人<br><span class="font-normal text-[10px]">財務核算簽字</span></span>
              </div>
              <div class="sig-space h-14 flex items-end justify-center text-slate-400 italic text-[9px] pb-1">
                (Ký và ghi rõ họ tên / 簽字蓋章)
              </div>
            </div>
            <div>
              <div class="font-bold text-slate-900 mb-1 uppercase">
                <span v-if="viewMode === 'customer'">XÁC NHẬN CỦA BÊN BÁN<br><span class="font-normal text-[10px]">賣方確認蓋章</span></span>
                <span v-else>GIÁM ĐỐC DUYỆT / 總經理審批<br><span class="font-normal text-[10px]">授權審批簽字</span></span>
              </div>
              <div class="sig-space h-14 flex flex-col justify-end items-center pb-1">
                <div class="font-bold text-slate-900 uppercase text-[10px]" v-text="companyInfo.name_vi"></div>
                <div class="font-semibold text-slate-800 uppercase mt-0.5 text-[10px]">GIÁM ĐỐC: WANG XIAN LI</div>
              </div>
            </div>
          </div>

          <!-- 頁尾版本與版權宣告 -->
          <div class="keep-together border-t border-slate-100 mt-3 pt-1.5 flex justify-between text-[8px] text-slate-400 font-mono">
            <span>GalaxyTeck Commercial System v3.5.0 · Trang 1/1</span>
            <span>Copyright © 2024-2026 GALAXY VIETNAM CO., LTD. All Rights Reserved.</span>
          </div>

        </div>

        <!-- ================= CRM 客戶管理中心彈窗 (支援客戶專屬訂單穿透查詢與調出維護) ================= -->
        <div v-show="showCustomerModal" style="display: none;" class="no-print fixed inset-0 bg-black/50 backdrop-blur-xs flex items-center justify-center p-4 z-50">
          <div class="bg-white rounded-2xl shadow-2xl max-w-5xl w-full p-6 max-h-[90vh] flex flex-col border border-slate-200">
            <div class="flex justify-between items-center border-b border-slate-100 pb-3 mb-3">
              <div class="flex items-center gap-2">
                <span class="text-base font-bold text-slate-800">👥 CRM 客戶檔案與訂單管理中心</span>
                <span class="text-xs text-slate-500">(共 {{ customerList.length }} 筆客戶檔案)</span>
              </div>
              <button type="button" @click="closeCustomerModal" class="text-slate-500 hover:text-slate-800 font-bold text-lg">✕</button>
            </div>

            <!-- 客戶名下訂單專屬檢視面板 (點擊某個客戶的訂單數時切換呈現) -->
            <div v-if="viewingCustomerQuotesFor" class="bg-blue-50/70 border border-blue-200 rounded-xl p-3.5 mb-3 space-y-2">
              <div class="flex justify-between items-center border-b border-blue-100 pb-2">
                <div class="flex items-center gap-2">
                  <span class="px-2 py-0.5 bg-blue-600 text-white rounded text-xs font-bold">歷史訂單</span>
                  <span class="font-bold text-slate-800 text-sm">【{{ viewingCustomerQuotesFor.company_name }}】名下的全部報價單/訂單</span>
                </div>
                <button type="button" @click="viewingCustomerQuotesFor = null" class="text-xs text-blue-600 hover:underline font-bold">
                  返回客戶總表 ✕
                </button>
              </div>

              <div class="overflow-y-auto max-h-56 text-xs bg-white rounded-lg border border-blue-100">
                <table class="w-full text-left border-collapse">
                  <thead class="bg-slate-50 border-b text-slate-500 sticky top-0">
                    <tr>
                      <th class="p-2.5">單號</th>
                      <th class="p-2.5">價格階梯</th>
                      <th class="p-2.5 text-right">含稅總金額</th>
                      <th class="p-2.5 text-center">狀態</th>
                      <th class="p-2.5 text-center">日期</th>
                      <th class="p-2.5 text-center w-28">操作</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr v-for="q in customerQuotesList" :key="q.id" class="border-b hover:bg-blue-50/50">
                      <td class="p-2.5 font-mono font-bold text-blue-600">{{ q.quote_no }}</td>
                      <td class="p-2.5">{{ q.customer_tier }}</td>
                      <td class="p-2.5 text-right font-mono font-bold">{{ formatCurrency(q.final_total_amount) }} {{ q.quote_currency }}</td>
                      <td class="p-2.5 text-center"><span class="px-2 py-0.5 bg-emerald-50 text-emerald-700 rounded text-[10px]">{{ q.status }}</span></td>
                      <td class="p-2.5 text-center text-slate-400 text-[11px]">{{ q.created_at }}</td>
                      <td class="p-2.5 text-center">
                        <button type="button" @click="loadQuoteDetailFromCRM(q.id)" class="px-3 py-1 bg-blue-600 hover:bg-blue-700 text-white rounded font-bold text-xs shadow-xs">
                          調出維護
                        </button>
                      </td>
                    </tr>
                    <tr v-if="customerQuotesList.length === 0">
                      <td colspan="6" class="p-6 text-center text-slate-400">該客戶目前尚無已存檔的報價單/訂單記錄</td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>

            <!-- 客戶新增/編輯表單 -->
            <div v-else class="bg-slate-50 p-4 border border-slate-200 rounded-xl mb-3 text-xs space-y-2.5">
              <div class="font-bold text-slate-800 flex justify-between">
                <span>{{ editingCustomer.id ? '✏ 編輯客戶資料' : '➕ 快速新增客戶檔案' }}</span>
                <button v-if="editingCustomer.id" type="button" @click="resetCustomerForm" class="text-[11px] text-blue-600 underline">切換為新增模式</button>
              </div>
              <div class="grid grid-cols-1 md:grid-cols-3 gap-2.5">
                <input v-model="editingCustomer.company_name" placeholder="* 客戶公司全稱" class="border border-slate-200 rounded-xl px-3 py-2 bg-white font-bold focus:border-blue-500 outline-none">
                <input v-model="editingCustomer.tax_code" placeholder="企業稅號 (Mã số thuế)" class="border border-slate-200 rounded-xl px-3 py-2 bg-white font-mono focus:border-blue-500 outline-none">
                <input v-model="editingCustomer.contact_person" placeholder="主要聯絡人" class="border border-slate-200 rounded-xl px-3 py-2 bg-white focus:border-blue-500 outline-none">
                <input v-model="editingCustomer.phone" placeholder="電話號碼" class="border border-slate-200 rounded-xl px-3 py-2 bg-white focus:border-blue-500 outline-none">
                <input v-model="editingCustomer.email" placeholder="電子郵件" class="border border-slate-200 rounded-xl px-3 py-2 bg-white focus:border-blue-500 outline-none">
                <select v-model="editingCustomer.default_tier" class="border border-slate-200 rounded-xl px-3 py-2 bg-white font-semibold text-slate-700 focus:border-blue-500 outline-none">
                  <option value="大型經銷商">預設階梯: 大型經銷商 (整櫃)</option>
                  <option value="中型經銷商">預設階梯: 中型經銷商 (>500萬)</option>
                  <option value="中型安裝商">預設階梯: 中型安裝商 (整托)</option>
                  <option value="小型安裝商">預設階梯: 小型安裝商 (散單)</option>
                </select>
                <input v-model="editingCustomer.address" placeholder="公司登記地址 / 倉庫送貨地址" class="md:col-span-2 border border-slate-200 rounded-xl px-3 py-2 bg-white focus:border-blue-500 outline-none">
                <div class="flex gap-2">
                  <button type="button" @click="submitSaveCustomer" class="flex-1 bg-blue-600 hover:bg-blue-700 text-white rounded-xl font-bold py-2 shadow-sm transition">
                    {{ editingCustomer.id ? '保存修改' : '立即新增建檔' }}
                  </button>
                  <button v-if="editingCustomer.id" type="button" @click="resetCustomerForm" class="px-3.5 bg-slate-200 text-slate-600 rounded-xl font-bold py-2">
                    取消
                  </button>
                </div>
              </div>
            </div>

            <!-- 客戶主列表 (包含訂單穿透徽章) -->
            <div class="overflow-y-auto flex-1 text-xs border border-slate-200 rounded-xl">
              <table class="w-full text-left border-collapse">
                <thead class="sticky top-0 bg-slate-50 text-slate-500 font-semibold border-b border-slate-200">
                  <tr>
                    <th class="p-3">客戶公司名稱</th>
                    <th class="p-3">稅號</th>
                    <th class="p-3">聯絡人 / 電話</th>
                    <th class="p-3">預設階梯</th>
                    <th class="p-3 text-center">訂單關聯</th>
                    <th class="p-3 text-center w-40">操作</th>
                  </tr>
                </thead>
                <tbody>
                  <tr v-for="c in customerList" :key="c.id" class="border-b border-slate-100 hover:bg-slate-50 transition">
                    <td class="p-3 font-bold text-slate-800">{{ c.company_name }}</td>
                    <td class="p-3 font-mono text-[11px] text-slate-500">{{ c.tax_code || '-' }}</td>
                    <td class="p-3">{{ c.contact_person }} <span class="text-slate-400">({{ c.phone || '-' }})</span></td>
                    <td class="p-3"><span class="px-2.5 py-0.5 bg-blue-50 text-blue-600 rounded-full text-[10px] font-semibold">{{ c.default_tier }}</span></td>
                    
                    <!-- 核心按鈕：查看該客戶名下的訂單 -->
                    <td class="p-3 text-center">
                      <button type="button" @click="openCustomerQuotesDrawer(c)" :class="['px-2.5 py-1 rounded-lg text-xs font-bold transition flex items-center gap-1 mx-auto', c.quote_count > 0 ? 'bg-sky-100 text-sky-800 hover:bg-sky-200 border border-sky-300' : 'bg-slate-100 text-slate-500 hover:bg-slate-200']">
                        <span>📜 訂單</span>
                        <span class="font-mono font-black text-[11px]">({{ c.quote_count || 0 }})</span>
                      </button>
                    </td>

                    <td class="p-3 text-center space-x-1">
                      <button type="button" @click="selectCustomerToForm(c); closeCustomerModal();" class="px-2.5 py-1 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-[11px] font-bold transition">
                        帶入報價單
                      </button>
                      <button type="button" @click="editCustomer(c)" class="px-2 py-1 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg text-[11px]">
                        改
                      </button>
                      <button type="button" @click="deleteCustomer(c.id)" class="px-2 py-1 bg-rose-50 hover:bg-rose-100 text-rose-600 rounded-lg text-[11px]">
                        刪
                      </button>
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>

            <div class="flex justify-end pt-3 border-t border-slate-100 mt-3">
              <button type="button" @click="closeCustomerModal" class="px-4 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl font-semibold text-xs">
                關閉
              </button>
            </div>
          </div>
        </div>

        <!-- 產品搜尋彈窗 -->
        <div v-show="showSearchModal" style="display: none;" class="no-print fixed inset-0 bg-black/50 backdrop-blur-xs flex items-center justify-center p-4 z-50">
          <div class="bg-white rounded-2xl shadow-2xl max-w-4xl w-full p-6 max-h-[85vh] flex flex-col border border-slate-200">
            <div class="flex justify-between items-center border-b border-slate-100 pb-3 mb-3">
              <div class="flex items-center gap-2">
                <span class="text-base font-bold text-slate-800">🔍 產品庫多維檢索與快捷選型</span>
                <span class="text-xs text-slate-500">(共有 {{ products.length }} 款商品)</span>
              </div>
              <button type="button" @click="showSearchModal = false" class="text-slate-500 hover:text-slate-800 font-bold text-lg">✕</button>
            </div>
            <div class="flex gap-2.5 mb-3">
              <input type="text" v-model="modalKeyword" ref="modalInput" placeholder="輸入型號、規格、扇葉直徑、功率或品牌..." class="flex-1 border border-slate-200 rounded-xl px-3.5 py-2 text-xs focus:ring-2 focus:ring-blue-100 focus:border-blue-500 outline-none font-medium">
              <select v-model="selectedCategoryFilter" class="border border-slate-200 rounded-xl px-3 py-2 text-xs bg-white text-slate-700 focus:border-blue-500 outline-none">
                <option value="">全部產品類別 (All)</option>
                <option v-for="cat in availableCategories" :key="cat" :value="cat">{{ cat }}</option>
              </select>
            </div>
            <div class="overflow-y-auto flex-1 text-xs border border-slate-200 rounded-xl">
              <table class="w-full text-left border-collapse">
                <thead class="sticky top-0 bg-slate-50 text-slate-500 font-semibold border-b border-slate-200">
                  <tr>
                    <th class="p-2.5 w-16">類別</th>
                    <th class="p-2.5">品名 / 型號</th>
                    <th class="p-2.5">詳細規格參數</th>
                    <th class="p-2.5 text-right w-28">基准底價({{ form.target_currency }})</th>
                    <th class="p-2.5 text-center w-20">操作</th>
                  </tr>
                </thead>
                <tbody>
                  <tr v-for="p in searchResults" :key="p.id" class="border-b border-slate-100 hover:bg-slate-50 transition">
                    <td class="p-2.5 text-[11px] text-slate-500">{{ p.category }}</td>
                    <td class="p-2.5 font-bold text-slate-800">{{ p.model }}</td>
                    <td class="p-2.5 text-slate-500 font-mono text-[10px]">{{ p.spec }}</td>
                    <td class="p-2.5 text-right font-mono font-bold text-slate-800">
                      {{ formatCurrency(getProductBasePriceInTargetCurrency(p)) }}
                    </td>
                    <td class="p-2.5 text-center">
                      <button type="button" @click="addSpecificProduct(p)" class="px-3 py-1 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-xs font-bold transition shadow-sm">
                        + 選入
                      </button>
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
            <div class="flex justify-end pt-3 border-t border-slate-100 mt-3">
              <button type="button" @click="showSearchModal = false" class="px-4 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl font-semibold text-xs">
                關閉
              </button>
            </div>
          </div>
        </div>

        <!-- 歷史單據庫彈窗 -->
        <div v-show="showHistoryModal" style="display: none;" class="no-print fixed inset-0 bg-black/50 backdrop-blur-xs flex items-center justify-center p-4 z-50">
          <div class="bg-white rounded-2xl shadow-2xl max-w-4xl w-full p-6 max-h-[80vh] flex flex-col border border-slate-200">
            <div class="flex justify-between items-center border-b border-slate-100 pb-3 mb-3">
              <div class="flex items-center gap-2">
                <h3 class="text-sm font-bold text-slate-800">報價歷史庫 (查詢、預覽與調出)</h3>
                <span class="text-xs text-blue-600 font-bold">(已歸檔 {{ historyList.length }} 份單據)</span>
              </div>
              <button type="button" @click="showHistoryModal = false" class="text-slate-500 hover:text-slate-800 font-bold text-lg">✕</button>
            </div>
            <div class="overflow-y-auto flex-1 text-xs border border-slate-200 rounded-xl">
              <table class="w-full text-left border-collapse">
                <thead>
                  <tr class="bg-slate-50 border-b border-slate-200 text-slate-500 font-semibold">
                    <th class="p-3 w-12 text-center">ID</th>
                    <th class="p-3">單號</th>
                    <th class="p-3">客戶名稱</th>
                    <th class="p-3 text-right">含稅總金額</th>
                    <th class="p-3 text-center">狀態</th>
                    <th class="p-3 text-center">存檔時間</th>
                    <th class="p-3 text-center w-28">操作</th>
                  </tr>
                </thead>
                <tbody>
                  <tr v-for="q in historyList" :key="q.id" class="border-b border-slate-100 hover:bg-slate-50">
                    <td class="p-3 text-center font-mono text-slate-400" v-text="q.id"></td>
                    <td class="p-3 font-mono font-bold text-blue-600" v-text="q.quote_no"></td>
                    <td class="p-3 font-medium text-slate-800" v-text="q.customer_name"></td>
                    <td class="p-3 text-right font-mono font-bold">
                      <span v-text="formatCurrency(q.final_total_amount)"></span> <span v-text="q.quote_currency"></span>
                    </td>
                    <td class="p-3 text-center">
                      <span class="px-2.5 py-0.5 bg-emerald-50 text-emerald-600 rounded-full text-[11px]" v-text="q.status"></span>
                    </td>
                    <td class="p-3 text-center text-slate-400 text-[10px]" v-text="q.created_at"></td>
                    <td class="p-3 text-center">
                      <button type="button" @click="loadQuoteDetail(q.id)" class="px-3 py-1 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-[11px] font-semibold transition shadow-xs">
                        調出維護
                      </button>
                    </td>
                  </tr>
                  <tr v-if="historyList.length === 0">
                    <td colspan="7" class="text-center py-8 text-slate-400">目前暫無歷史單據記錄</td>
                  </tr>
                </tbody>
              </table>
            </div>
            <div class="flex justify-end pt-3 border-t border-slate-100 mt-3">
              <button type="button" @click="showHistoryModal = false" class="px-4 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl font-semibold text-xs">
                關閉
              </button>
            </div>
          </div>
        </div>

      </div>

      <script>
        const { createApp } = Vue;

        createApp({
          data() {
            return {
              viewMode: 'customer',
              printDensity: 'auto',
              products: [],
              rates: {},
              selectedProductIdToAdd: null,
              currentQuoteId: null,
              todayDate: '',
              validUntilDate: '',
              showTemplateDrawer: false,
              showHistoryModal: false,
              showSearchModal: false,
              showCustomerModal: false,
              showCustomerDropdown: false,
              
              // CRM 訂單穿透專用狀態
              viewingCustomerQuotesFor: null,
              customerQuotesList: [],

              layoutTemplatesList: [],
              selectedLayoutTemplateId: 1,

              customerList: [],
              editingCustomer: {
                id: null, company_name: '', short_name: '', tax_code: '',
                contact_person: '', phone: '', email: '', address: '',
                default_tier: '中型經銷商', payment_terms: '', notes: ''
              },
              quickSearchKeyword: '',
              modalKeyword: '',
              selectedCategoryFilter: '',
              historyList: [],

              companyInfo: {
                logo_data: null,
                name_vi: 'CÔNG TY TNHH TẬP ĐOÀN CÔNG NGHỆ GALAXY VIỆT NAM',
                name_cn: '銀河科技集團（越南）有限公司',
                tax_code: '2301296587',
                contact_info: 'Điện thoại: 0325.609.218 | Email: galaxyteck6886@gmail.com',
                address: 'Số 36 Yết Kiêu, Phường Kinh Bắc, Tỉnh Bắc Ninh, Việt Nam',
                bank_info: 'STK (Techcombank): 314431 - Ngân hàng TMCP Kỹ thương Việt Nam'
              },

              layoutConfig: {
                title_text: 'BẢNG BÁO GIÁ 商業報價單',
                greeting_text: 'Chúng tôi xin gửi đến Quý khách hàng bảng báo giá như sau / 我司現向貴司呈報以下報價表:'
              },

              form: {
                quote_no: '',
                layout_template_id: 1,
                customer_id: 0,
                customer_name: 'CÔNG TY TNHH CƠ ĐIỆN VÀ NĂNG LƯỢNG BẮC NINH',
                contact_person: 'Nguyễn Văn Nam',
                phone_email: '0988.123.456',
                sales_rep: 'Thanh Bình 清平',
                customer_tier: '中型經銷商',
                delivery_scenario: '越南本地倉',
                trade_term: 'DAP',
                target_currency: 'VND',
                vat_rate_pct: 10.0,
                valid_days: 3,
                markup_pct: 0.0,
                lang_mode: 'vi_zh',
                template_style: 'standard'
              },

              termsList: [
                'Báo giá có hiệu lực trong vòng 3 ngày kể từ ngày phát hành / 本報價有效期為 3 天。',
                'Thời gian giao hàng: 3-5 ngày sau khi đặt hàng / 交貨期：下單後 3-5 天內安排發貨。',
                'Giá sản phẩm chưa bao gồm chi phí vận chuyển và lắp đặt / 單價不包含現場運輸及安裝費用。',
                'Địa điểm giao hàng: Giao tại kho khách hàng chỉ định / 交貨地點：客戶指定工廠倉庫。',
                'Thanh toán: Chuyển khoản, đặt cọc 50% khi xác nhận đơn, thanh toán 50% trước khi giao hàng / 付款方式：銀行電匯轉賬，確認訂單預付 50%，發貨前付清剩餘 50%。'
              ],
              tableRows: []
            }
          },
          computed: {
            matchedCustomers() {
              if (!this.form.customer_name) return this.customerList.slice(0, 5);
              const kw = this.form.customer_name.trim().toLowerCase();
              return this.customerList.filter(c => 
                c.company_name.toLowerCase().includes(kw) || 
                (c.short_name && c.short_name.toLowerCase().includes(kw))
              ).slice(0, 6);
            },
            hasAnyCustomMarkup() {
              return this.tableRows.some(r => r.custom_markup !== null && r.custom_markup !== undefined && r.custom_markup !== '');
            },
            availableCategories() {
              const cats = new Set();
              this.products.forEach(p => { if (p.category) cats.add(p.category); });
              return Array.from(cats);
            },
            searchResults() {
              return this.products.filter(p => {
                const kw = this.modalKeyword.trim().toLowerCase();
                const matchKw = !kw || 
                  (p.model && p.model.toLowerCase().includes(kw)) ||
                  (p.spec && p.spec.toLowerCase().includes(kw)) ||
                  (p.brand && p.brand.toLowerCase().includes(kw)) ||
                  (p.category && p.category.toLowerCase().includes(kw));
                const matchCat = !this.selectedCategoryFilter || p.category === this.selectedCategoryFilter;
                return matchKw && matchCat;
              });
            },
            filteredProductList() {
              if (!this.quickSearchKeyword) return this.products;
              const kw = this.quickSearchKeyword.trim().toLowerCase();
              return this.products.filter(p => {
                return (p.model && p.model.toLowerCase().includes(kw)) ||
                       (p.spec && p.spec.toLowerCase().includes(kw)) ||
                       (p.brand && p.brand.toLowerCase().includes(kw));
              });
            },
            globalMarkupFactor() {
              const pct = Number(this.form.markup_pct);
              if (isNaN(pct)) return 1.0;
              return 1.0 + (pct / 100.0);
            },
            totalBeforeTax() {
              return this.tableRows.reduce((sum, r) => sum + this.calcRowSubtotal(r), 0);
            },
            vatAmount() {
              const vat = this.totalBeforeTax * (Number(this.form.vat_rate_pct || 0) / 100.0);
              return this.form.target_currency === 'VND' ? Math.round(vat) : Number(vat.toFixed(2));
            },
            finalTotalAfterTax() {
              return this.totalBeforeTax + this.vatAmount;
            },
            totalBaseCost() {
              return this.tableRows.reduce((sum, r) => {
                const q = Number(r.quantity) || 0;
                const b = Number(r.base_price) || 0;
                return sum + (q * b);
              }, 0);
            },
            totalGrossProfit() {
              return this.totalBeforeTax - this.totalBaseCost;
            },
            grossMarginPercent() {
              if (this.totalBeforeTax <= 0) return 0.0;
              return (this.totalGrossProfit / this.totalBeforeTax) * 100.0;
            }
          },
          async mounted() {
            this.initDateAndNo();
            await this.loadInitialData();
            await this.loadGlobalSettings();
            await this.loadLayoutTemplatesList();
            await this.loadCustomerList();
            await this.loadHistoryListSilent();
            document.addEventListener('click', (e) => {
              if (!e.target.closest('.relative')) {
                this.showCustomerDropdown = false;
              }
            });
          },
          methods: {
            async fetchNextQuoteNo() {
              try {
                const res = await fetch('/api/quotes/generate_next_no');
                const data = await res.json();
                if (data.next_quote_no) {
                  this.form.quote_no = data.next_quote_no;
                }
              } catch (e) {
                console.error("領取單號失敗:", e);
              }
            },

            async handleSafeCreateNewQuote() {
              if (this.tableRows.length > 0 || this.form.customer_name) {
                if (!confirm("⚠ 確認要新建空白報價單嗎？\\n當前畫布內容將被重置，系統將為您生成下一筆全新單號。")) {
                  return;
                }
              }
              this.createNewQuote();
              await this.fetchNextQuoteNo();
            },

            async handleLogoUploadAndPersist(e) {
              const file = e.target.files[0];
              if (!file) return;
              if (file.size > 2 * 1024 * 1024) {
                alert("圖片大小請控制在 2MB 以內！");
                return;
              }
              const reader = new FileReader();
              reader.onload = async (uploadEvent) => {
                const b64 = uploadEvent.target.result;
                this.companyInfo.logo_data = b64;
                try {
                  const res = await fetch('/api/system/save_global_logo', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ logo_data: b64 })
                  });
                  await res.json();
                } catch (err) {
                  console.error("保存全局 Logo 失敗:", err);
                }
              };
              reader.readAsDataURL(file);
            },
            async loadGlobalSettings() {
              try {
                const res = await fetch('/api/system/global_settings');
                const data = await res.json();
                if (data && data.logo_data) {
                  this.companyInfo.logo_data = data.logo_data;
                }
              } catch (e) {
                console.error("載入全局 Logo 失敗:", e);
              }
            },
            async loadLayoutTemplatesList() {
              try {
                const res = await fetch('/api/layout_templates/list');
                this.layoutTemplatesList = await res.json();
                if (this.layoutTemplatesList.length > 0) {
                  const defTpl = this.layoutTemplatesList.find(t => t.is_default == 1) || this.layoutTemplatesList[0];
                  this.selectedLayoutTemplateId = defTpl.id;
                  if (!this.companyInfo.logo_data && defTpl.logo_data) {
                    this.companyInfo.logo_data = defTpl.logo_data;
                  }
                }
              } catch (e) {
                console.error("載入版式範本庫失敗:", e);
              }
            },
            applyLayoutTemplateById(tplId) {
              const tpl = this.layoutTemplatesList.find(t => t.id == tplId);
              if (!tpl) return;
              if (tpl.logo_data) this.companyInfo.logo_data = tpl.logo_data;
              this.companyInfo.name_vi = tpl.name_vi || '';
              this.companyInfo.name_cn = tpl.name_cn || '';
              this.companyInfo.tax_code = tpl.tax_code || '';
              this.companyInfo.contact_info = tpl.contact_info || '';
              this.companyInfo.address = tpl.address || '';
              this.companyInfo.bank_info = tpl.bank_info || '';
              this.layoutConfig.title_text = tpl.title_text || 'BẢNG BÁO GIÁ 商業報價單';
              this.layoutConfig.greeting_text = tpl.greeting_text || '';
              this.form.lang_mode = tpl.lang_mode || 'vi_zh';
              this.form.layout_template_id = tpl.id;

              if (tpl.terms_list && tpl.terms_list.length > 0) {
                this.termsList = [...tpl.terms_list];
              }
            },
            async promptSaveNewLayoutTemplate() {
              const name = prompt("請輸入新版式範本名稱 (例如: 光伏項目合約專用版式 / 外貿英文出口版):");
              if (!name || !name.trim()) return;

              const payload = {
                template_name: name.trim(),
                is_default: 0,
                logo_data: this.companyInfo.logo_data,
                name_vi: this.companyInfo.name_vi,
                name_cn: this.companyInfo.name_cn,
                tax_code: this.companyInfo.tax_code,
                contact_info: this.companyInfo.contact_info,
                address: this.companyInfo.address,
                bank_info: this.companyInfo.bank_info,
                title_text: this.layoutConfig.title_text,
                greeting_text: this.layoutConfig.greeting_text,
                lang_mode: this.form.lang_mode,
                template_style: this.form.template_style,
                terms_list: this.termsList
              };

              const res = await fetch('/api/layout_templates/save', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
              });
              const data = await res.json();
              if (data.status === 'success') {
                alert("🎉 新版式範本已成功保存！");
                await this.loadLayoutTemplatesList();
                this.selectedLayoutTemplateId = data.id;
              }
            },
            async updateCurrentLayoutTemplate() {
              if (!this.selectedLayoutTemplateId) return;
              const cur = this.layoutTemplatesList.find(t => t.id == this.selectedLayoutTemplateId);
              if (!confirm(`確定要覆蓋更新當前版式範本「${cur ? cur.template_name : ''}」嗎？`)) return;

              const payload = {
                id: this.selectedLayoutTemplateId,
                template_name: cur ? cur.template_name : '自定義版式',
                is_default: cur ? cur.is_default : 0,
                logo_data: this.companyInfo.logo_data,
                name_vi: this.companyInfo.name_vi,
                name_cn: this.companyInfo.name_cn,
                tax_code: this.companyInfo.tax_code,
                contact_info: this.companyInfo.contact_info,
                address: this.companyInfo.address,
                bank_info: this.companyInfo.bank_info,
                title_text: this.layoutConfig.title_text,
                greeting_text: this.layoutConfig.greeting_text,
                lang_mode: this.form.lang_mode,
                template_style: this.form.template_style,
                terms_list: this.termsList
              };

              const res = await fetch('/api/layout_templates/save', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
              });
              const data = await res.json();
              if (data.status === 'success') {
                alert("✅ 當前版式範本更新成功！");
                await this.loadLayoutTemplatesList();
              }
            },
            async deleteCurrentLayoutTemplate() {
              if (!this.selectedLayoutTemplateId) return;
              const cur = this.layoutTemplatesList.find(t => t.id == this.selectedLayoutTemplateId);
              if (!confirm(`警告：確定要刪除版式範本「${cur ? cur.template_name : ''}」嗎？此操作不可逆！`)) return;

              await fetch(`/api/layout_templates/${this.selectedLayoutTemplateId}`, { method: 'DELETE' });
              alert("已成功刪除該版式範本！");
              await this.loadLayoutTemplatesList();
              if (this.layoutTemplatesList.length > 0) {
                this.applyLayoutTemplateById(this.layoutTemplatesList[0].id);
              }
            },

            // ================= CRM 客戶管理與訂單穿透查詢 =================
            async loadCustomerList() {
              try {
                const res = await fetch('/api/customers/list');
                this.customerList = await res.json();
              } catch (e) {
                console.error("載入客戶列表失敗:", e);
              }
            },
            openCustomerModal() {
              this.showCustomerModal = true;
              this.viewingCustomerQuotesFor = null;
              this.loadCustomerList();
            },
            closeCustomerModal() {
              this.showCustomerModal = false;
              this.viewingCustomerQuotesFor = null;
            },
            async openCustomerQuotesDrawer(c) {
              this.viewingCustomerQuotesFor = c;
              this.customerQuotesList = [];
              try {
                const res = await fetch(`/api/customers/${c.id}/quotes`);
                this.customerQuotesList = await res.json();
              } catch (e) {
                console.error("載入客戶訂單歷史失敗:", e);
              }
            },
            async loadQuoteDetailFromCRM(id) {
              this.closeCustomerModal();
              await this.loadQuoteDetail(id);
            },
            selectCustomerToForm(c) {
              this.form.customer_id = c.id;
              this.form.customer_name = c.company_name;
              this.form.contact_person = c.contact_person || '';
              this.form.phone_email = c.phone ? `${c.phone}${c.email ? ' / ' + c.email : ''}` : (c.email || '');
              if (c.default_tier) this.form.customer_tier = c.default_tier;
              this.showCustomerDropdown = false;
            },
            editCustomer(c) {
              this.viewingCustomerQuotesFor = null;
              this.editingCustomer = { ...c };
            },
            resetCustomerForm() {
              this.editingCustomer = {
                id: null, company_name: '', short_name: '', tax_code: '',
                contact_person: '', phone: '', email: '', address: '',
                default_tier: '中型經銷商', payment_terms: '', notes: ''
              };
            },
            async submitSaveCustomer() {
              if (!this.editingCustomer.company_name) {
                alert("請填寫客戶公司全稱！");
                return;
              }
              const res = await fetch('/api/customers/save', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(this.editingCustomer)
              });
              const data = await res.json();
              if (data.status === 'success') {
                alert("✅ 客戶檔案儲存成功！");
                this.resetCustomerForm();
                await this.loadCustomerList();
              }
            },
            async deleteCustomer(id) {
              if (!confirm("確定要刪除該客戶檔案嗎？")) return;
              await fetch(`/api/customers/${id}`, { method: 'DELETE' });
              await this.loadCustomerList();
            },

            // 結構化條款引擎
            addTermRow() {
              this.termsList.push('新增商務約定條款內容...');
            },
            removeTermRow(idx) {
              this.termsList.splice(idx, 1);
            },
            moveTermUp(idx) {
              if (idx > 0) {
                const temp = this.termsList[idx];
                this.termsList[idx] = this.termsList[idx - 1];
                this.termsList[idx - 1] = temp;
              }
            },
            moveTermDown(idx) {
              if (idx < this.termsList.length - 1) {
                const temp = this.termsList[idx];
                this.termsList[idx] = this.termsList[idx + 1];
                this.termsList[idx + 1] = temp;
              }
            },
            getEffectiveDensityClass() {
              if (this.printDensity === 'ultra') return 'compact-mode';
              if (this.printDensity === 'normal') return '';
              return this.tableRows.length <= 6 ? 'compact-mode' : '';
            },
            openSearchModal() {
              this.showSearchModal = true;
              this.modalKeyword = this.quickSearchKeyword;
              this.$nextTick(() => {
                if (this.$refs.modalInput) this.$refs.modalInput.focus();
              });
            },
            quickAddFirstMatch() {
              if (this.filteredProductList.length > 0) {
                this.addSpecificProduct(this.filteredProductList[0]);
                this.quickSearchKeyword = '';
              }
            },
            resetAllCustomMarkups() {
              this.tableRows.forEach(r => { r.custom_markup = null; });
            },
            getProductBasePriceInTargetCurrency(p) {
              const rateVnd = this.rates['USD_VND'] || 25967.0;
              const rateCny = this.rates['USD_CNY'] || 6.70;
              let baseP = p.default_base_usd || 100;
              if (this.form.target_currency === 'VND') return Math.round(baseP * rateVnd);
              if (this.form.target_currency === 'CNY') return Number((baseP * rateCny).toFixed(2));
              return Number(baseP.toFixed(2));
            },
            addSpecificProduct(p) {
              let unitStr = '台 / Bộ';
              if (p.category && p.category.includes('組件')) unitStr = '塊 / Tấm';
              else if (p.category && p.category.includes('逆變器')) unitStr = '套 / Bộ';

              const baseP = this.getProductBasePriceInTargetCurrency(p);
              this.tableRows.push({
                product_id: p.id,
                name: p.model,
                spec: p.spec || p.model,
                unit: unitStr,
                quantity: 1,
                base_price: baseP,
                custom_markup: null,
                note: ''
              });
            },
            getRowEffectiveMarkupFactor(row) {
              if (row.custom_markup !== null && row.custom_markup !== undefined && row.custom_markup !== '') {
                const rowPct = Number(row.custom_markup);
                return isNaN(rowPct) ? 1.0 : (1.0 + rowPct / 100.0);
              }
              return this.globalMarkupFactor;
            },
            calcRowFinalUnitPrice(row) {
              const base = Number(row.base_price) || 0;
              const factor = this.getRowEffectiveMarkupFactor(row);
              const finalPrice = base * factor;
              if (this.form.target_currency === 'VND') {
                return Math.round(finalPrice);
              } else {
                return Number(finalPrice.toFixed(2));
              }
            },
            onManualUnitPriceChange(row, newPriceVal) {
              const newP = Number(newPriceVal) || 0;
              if (row.product_id === 0) {
                if (row.custom_markup === null || row.custom_markup === undefined || row.custom_markup === '') {
                  const factor = this.globalMarkupFactor > 0 ? this.globalMarkupFactor : 1.0;
                  const calculatedBase = newP / factor;
                  row.base_price = this.form.target_currency === 'VND' ? Math.round(calculatedBase) : Number(calculatedBase.toFixed(2));
                } else {
                  const factor = this.getRowEffectiveMarkupFactor(row);
                  const calculatedBase = factor > 0 ? (newP / factor) : newP;
                  row.base_price = this.form.target_currency === 'VND' ? Math.round(calculatedBase) : Number(calculatedBase.toFixed(2));
                }
                return;
              }

              const baseP = Number(row.base_price) || 0;
              if (baseP > 0) {
                const ratio = ((newP - baseP) / baseP) * 100.0;
                row.custom_markup = Number(ratio.toFixed(2));
              } else {
                row.base_price = newP;
                row.custom_markup = null;
              }
            },
            calcRowSubtotal(row) {
              const q = Number(row.quantity) || 0;
              const unitP = this.calcRowFinalUnitPrice(row);
              const sub = q * unitP;
              if (this.form.target_currency === 'VND') {
                return Math.round(sub);
              } else {
                return Number(sub.toFixed(2));
              }
            },
            calcRowProfit(row) {
              const q = Number(row.quantity) || 0;
              const unitP = this.calcRowFinalUnitPrice(row);
              const baseP = Number(row.base_price) || 0;
              const profit = q * (unitP - baseP);
              if (this.form.target_currency === 'VND') {
                return Math.round(profit);
              } else {
                return Number(profit.toFixed(2));
              }
            },
            initDateAndNo() {
              const now = new Date();
              const y = now.getFullYear();
              const m = String(now.getMonth() + 1).padStart(2, '0');
              const d = String(now.getDate()).padStart(2, '0');
              this.todayDate = `${d}.${m}.${y}`;
              if (!this.form.quote_no) {
                this.form.quote_no = `BG-${y}.${m}.${d}-001`;
              }

              const validDate = new Date();
              validDate.setDate(now.getDate() + this.form.valid_days);
              const vy = validDate.getFullYear();
              const vm = String(validDate.getMonth() + 1).padStart(2, '0');
              const vd = String(validDate.getDate()).padStart(2, '0');
              this.validUntilDate = `${vd}.${vm}.${vy}`;
            },
            async loadInitialData() {
              try {
                const [pRes, rRes] = await Promise.all([
                  fetch('/api/products'),
                  fetch('/api/rates')
                ]);
                this.products = await pRes.json();
                this.rates = await rRes.json();
                
                if (this.products.length > 0 && this.tableRows.length === 0) {
                  this.addSpecificProduct(this.products[0]);
                }
              } catch (e) {
                console.error("載入失敗:", e);
              }
            },
            onCurrencyChange() {
              this.tableRows.forEach(r => {
                if (r.product_id > 0) {
                  const p = this.products.find(x => x.id === r.product_id);
                  if (p) r.base_price = this.getProductBasePriceInTargetCurrency(p);
                }
              });
            },
            addCustomBlankRow() {
              this.tableRows.push({
                product_id: 0,
                name: '自定義產品品名 / Description',
                spec: '規格參數 / Spec',
                unit: '台 / 個 / pcs',
                quantity: 1,
                base_price: 0,
                custom_markup: null,
                note: ''
              });
            },
            removeRow(idx) {
              this.tableRows.splice(idx, 1);
            },

            // 儲存報價單 (支援覆蓋更新 vs 另存為新單據)
            async saveToServer(forceAsNew = false) {
              if (this.tableRows.length === 0) {
                alert("⚠️ 請先添加物料明細再保存報價單！");
                return;
              }

              const payload = {
                quote_id: forceAsNew ? null : this.currentQuoteId,
                save_as_new: forceAsNew,
                layout_template_id: this.selectedLayoutTemplateId || 1,
                customer_id: this.form.customer_id,
                ...this.form,
                custom_title: this.layoutConfig.title_text,
                custom_greeting: this.layoutConfig.greeting_text,
                terms_text: this.termsList.join('\\n'),
                terms_list: this.termsList,
                
                company_logo_data: this.companyInfo.logo_data,
                company_name_vi: this.companyInfo.name_vi,
                company_name_cn: this.companyInfo.name_cn,
                company_tax_code: this.companyInfo.tax_code,
                company_contact_info: this.companyInfo.contact_info,
                company_address: this.companyInfo.address,
                company_bank_info: this.companyInfo.bank_info,

                total_before_tax: this.totalBeforeTax,
                vat_amount: this.vatAmount,
                final_total_after_tax: this.finalTotalAfterTax,
                items: this.tableRows.map(r => ({
                  product_id: r.product_id || 0,
                  name: r.name,
                  spec: r.spec,
                  unit: r.unit,
                  quantity: Number(r.quantity),
                  base_price: Number(r.base_price),
                  custom_markup: (r.custom_markup !== null && r.custom_markup !== undefined && r.custom_markup !== '') ? Number(r.custom_markup) : null,
                  unit_price: this.calcRowFinalUnitPrice(r),
                  subtotal: this.calcRowSubtotal(r),
                  note: r.note || ''
                }))
              };

              try {
                const res = await fetch('/api/quote/save_or_update', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/json' },
                  body: JSON.stringify(payload)
                });
                const data = await res.json();
                
                if (res.ok && data.status === 'success') {
                  this.currentQuoteId = data.quote_id;
                  this.form.quote_no = data.quote_no;
                  alert("🎉 " + data.message);
                  await this.loadInitialData();
                  await this.loadCustomerList(); // 更新客戶訂單計數
                  await this.loadHistoryListSilent();
                } else {
                  alert("❌ 保存失敗: " + (data.detail || data.message || "資料庫服務異常"));
                }
              } catch (networkErr) {
                alert("❌ 保存請求失敗: 請檢查網路連線或後端服務是否正在運行！\\n" + networkErr);
              }
            },

            async loadHistoryListSilent() {
              try {
                const res = await fetch('/api/quotes/list');
                this.historyList = await res.json();
              } catch (e) {
                this.historyList = [];
              }
            },
            async openHistoryModal() {
              this.showHistoryModal = true;
              await this.loadHistoryListSilent();
            },
            async loadQuoteDetail(id) {
              try {
                const res = await fetch(`/api/quotes/${id}`);
                const data = await res.json();
                const h = data.header;
                this.currentQuoteId = h.id;
                this.form.quote_no = h.quote_no;
                this.form.customer_id = h.customer_id || 0;
                this.form.customer_name = h.customer_name;
                this.form.contact_person = h.contact_person || '';
                this.form.phone_email = h.phone_email || '';
                this.form.sales_rep = h.sales_rep || 'Thanh Bình 清平';
                this.form.customer_tier = h.customer_tier;
                this.form.target_currency = h.quote_currency;
                this.form.vat_rate_pct = Number(h.tax_rate);
                this.form.markup_pct = Number(h.markup_rate);
                this.form.valid_days = h.valid_days || 3;
                this.form.lang_mode = h.lang_mode || 'vi_zh';

                if (h.layout_template_id) this.selectedLayoutTemplateId = h.layout_template_id;

                if (h.terms_json) {
                  try {
                    this.termsList = JSON.parse(h.terms_json);
                  } catch (e) {
                    this.termsList = h.terms_text ? h.terms_text.split('\\n') : [];
                  }
                } else if (h.terms_text) {
                  this.termsList = h.terms_text.split('\\n');
                }

                if (h.custom_title) this.layoutConfig.title_text = h.custom_title;
                if (h.custom_greeting) this.layoutConfig.greeting_text = h.custom_greeting;

                if (h.company_logo_data) this.companyInfo.logo_data = h.company_logo_data;
                if (h.company_name_vi) this.companyInfo.name_vi = h.company_name_vi;
                if (h.company_name_cn) this.companyInfo.name_cn = h.company_name_cn;
                if (h.company_tax_code) this.companyInfo.tax_code = h.company_tax_code;
                if (h.company_contact_info) this.companyInfo.contact_info = h.company_contact_info;
                if (h.company_address) this.companyInfo.address = h.company_address;
                if (h.company_bank_info) this.companyInfo.bank_info = h.company_bank_info;

                this.tableRows = data.items.map(it => ({
                  product_id: it.product_id,
                  name: it.item_name || it.item_model,
                  spec: it.item_spec || '',
                  unit: it.unit || '台 / 套',
                  quantity: it.quantity,
                  base_price: Number(it.base_price) || Number(it.quote_unit_price),
                  custom_markup: it.custom_markup !== null ? Number(it.custom_markup) : null,
                  note: it.note || ''
                }));

                this.showHistoryModal = false;
                alert(`✅ 已成功調出報價單: ${h.quote_no} (ID: ${h.id})！\\n若修改後需要建立新單據，請點選「另存為新報價單」按鈕。`);
              } catch (loadErr) {
                alert("❌ 加載歷史報價單失敗: " + loadErr);
              }
            },
            createNewQuote() {
              this.currentQuoteId = null;
              this.form.customer_id = 0;
              this.form.customer_name = '';
              this.form.contact_person = '';
              this.form.phone_email = '';
              this.form.markup_pct = 0.0;
              this.form.quote_no = '';
              this.initDateAndNo();
              this.tableRows = [];
              if (this.products.length > 0) this.addSpecificProduct(this.products[0]);
            },
            printOffer() {
              window.print();
            },
            formatCurrency(val) {
              if (val === null || val === undefined || isNaN(val)) return '0';
              const num = Number(val);
              if (this.form.target_currency === 'VND') {
                return num.toLocaleString('vi-VN', { maximumFractionDigits: 0 });
              } else {
                return num.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
              }
            }
          }
        }).mount('#app');
      </script>
    </body>
    </html>
    """

if __name__ == '__main__':
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)