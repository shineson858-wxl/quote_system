# -*- coding: utf-8 -*-
"""
==================================================================================================
項目名稱 (Project): GALAXYTECK 智慧快速報價與客戶關係管理系統 (企業標準網絡版)
版本編號 (Version): v4.0.2 Gemini 3.8 Flash Edition with Auto-Retry
發布日期 (Date): 2026年10月
軟體授權與智慧財產權聲明 (Intellectual Property & Copyright):
    版權所有 (C) 2024-2026 CÔNG TY TNHH TẬP ĐOÀN CÔNG NGHỆ GALAXY VIỆT NAM
    (GALAXY VIETNAM TECHNOLOGY CORPORATION COMPANY LIMITED / 銀河科技集團(越南)有限公司)
    企業代碼 / Mã số doanh nghiệp: 2301296587
    地址 / Địa chỉ: Số 36 Yết Kiêu, Phường Kinh Bắc, Tỉnh Bắc Ninh, Việt Nam
==================================================================================================
"""

from fastapi import FastAPI, HTTPException, UploadFile, File
from fastapi.responses import HTMLResponse
from pydantic import BaseModel
from typing import List, Optional, Any, Dict
from decimal import Decimal, ROUND_HALF_UP
import pymysql
import datetime
import uvicorn
import pandas as pd
import io
import json
import uuid
import re
import os
import hashlib
import mimetypes
import time as _time
import random as _random

# ================= Gemini AI 配置（新版 SDK） =================
try:
    from google import genai
    from google.genai import types as genai_types
    GEMINI_AVAILABLE = True
except ImportError:
    GEMINI_AVAILABLE = False
    print("⚠️ google-genai 未安裝，請執行：pip install google-genai pillow")

GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
if not GEMINI_API_KEY:
    GEMINI_API_KEY = ""  # ← 也可在此直接填 Key

def _sanitize_api_key(raw: str) -> str:
    """API Key 會被放進 HTTP 請求頭，請求頭只允許 ASCII。
    這裡去掉首尾空白/引號，並剔除中文、全角字符、零寬字符等非 ASCII 內容，
    避免出現 'ascii' codec can't encode characters 錯誤。"""
    if not raw:
        return ""
    key = raw.strip().strip('"').strip("'").strip()
    cleaned = "".join(ch for ch in key if ch.isascii() and ch.isprintable() and not ch.isspace())
    if cleaned != key:
        print("⚠️ GEMINI_API_KEY 中含有非 ASCII 字符（如中文/全角符號/空格），已自動清除；"
              "請檢查環境變數或代碼中的 Key 是否為純英文數字的真實密鑰。")
    return cleaned

GEMINI_API_KEY = _sanitize_api_key(GEMINI_API_KEY)

# ✅ 您的專屬模型
GEMINI_MODEL = "gemini-3.8-flash"

# 新版 SDK 用 Client 實例
gemini_client = None
if GEMINI_AVAILABLE and GEMINI_API_KEY:
    try:
        if GEMINI_API_KEY.startswith("AQ."):
            # Vertex AI Express Mode 的 Key（以 AQ. 開頭）需要走 Vertex 接口
            gemini_client = genai.Client(vertexai=True, api_key=GEMINI_API_KEY)
        else:
            gemini_client = genai.Client(api_key=GEMINI_API_KEY)
        print(f"✅ Gemini AI 已啟用（新版 SDK），模型：{GEMINI_MODEL}")
    except Exception as e:
        print(f"⚠️ Gemini 初始化失敗：{e}")
        GEMINI_AVAILABLE = False
else:
    if GEMINI_AVAILABLE:
        print("⚠️ 未設置 GEMINI_API_KEY，AI 功能不可用（Excel 導入仍可正常使用）")

app = FastAPI(
    title="GalaxyTeck 智慧快速報價系統",
    version="4.0.2",
    description="企業級專業報價軟體：Gemini 3.8 AI 多模態導入、三語切換、CRM 訂單穿透、A4防跑版"
)

# ================= 1. 資料庫連線 =================
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

# ================= 2. 自愈引擎 =================
@app.on_event("startup")
def self_healing_db():
    try:
        conn = get_db()
        with conn.cursor() as cursor:
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

            cursor.execute("""
                CREATE TABLE IF NOT EXISTS `tb_product` (
                    `id` INT AUTO_INCREMENT PRIMARY KEY,
                    `sku` VARCHAR(100) DEFAULT '',
                    `brand` VARCHAR(100) DEFAULT '',
                    `category` VARCHAR(100) DEFAULT '',
                    `model` VARCHAR(150) DEFAULT '',
                    `spec` VARCHAR(255) DEFAULT '',
                    `power_w` INT DEFAULT 0,
                    `pcs_per_pallet` INT DEFAULT 1,
                    `pcs_per_container` INT DEFAULT 1,
                    `warranty_years` INT DEFAULT 1,
                    `is_active` TINYINT DEFAULT 1,
                    `created_at` TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
            """)

            cursor.execute("""
                CREATE TABLE IF NOT EXISTS `tb_price_matrix` (
                    `id` INT AUTO_INCREMENT PRIMARY KEY,
                    `product_id` INT NOT NULL,
                    `region` VARCHAR(50) DEFAULT '',
                    `trade_term` VARCHAR(20) DEFAULT 'DAP',
                    `customer_tier` VARCHAR(50) DEFAULT '',
                    `currency` VARCHAR(10) DEFAULT 'USD',
                    `unit_price` DECIMAL(15,2) DEFAULT 0.00,
                    UNIQUE KEY `uk_matrix` (`product_id`, `region`, `trade_term`, `customer_tier`, `currency`)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
            """)

            cursor.execute("""
                CREATE TABLE IF NOT EXISTS `tb_exchange_rate` (
                    `pair` VARCHAR(20) PRIMARY KEY,
                    `rate` DECIMAL(15,4) DEFAULT 1.0,
                    `updated_at` TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
            """)
            cursor.execute("""
                INSERT INTO `tb_exchange_rate` (`pair`, `rate`) VALUES
                ('USD_VND', 25967.0000), ('USD_CNY', 6.7000), ('CNY_VND', 3874.0000),
                ('USD_THB', 32.0500), ('USD_MYR', 3.9855)
                ON DUPLICATE KEY UPDATE rate=VALUES(rate);
            """)

            cursor.execute("""
                CREATE TABLE IF NOT EXISTS `tb_import_log` (
                    `id` INT AUTO_INCREMENT PRIMARY KEY,
                    `file_name` VARCHAR(255) DEFAULT '',
                    `import_mode` VARCHAR(50) DEFAULT 'library',
                    `import_source` VARCHAR(30) DEFAULT 'excel',
                    `ai_model` VARCHAR(50) DEFAULT '',
                    `ai_confidence` DECIMAL(5,2) DEFAULT 0.00,
                    `customer_id` INT DEFAULT 0,
                    `customer_name` VARCHAR(255) DEFAULT '',
                    `total_products` INT DEFAULT 0,
                    `new_products` INT DEFAULT 0,
                    `existing_products` INT DEFAULT 0,
                    `quotation_id` INT DEFAULT 0,
                    `raw_data_json` LONGTEXT DEFAULT NULL,
                    `created_at` TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
            """)
            for col, col_type in [
                ("import_source", "VARCHAR(30) DEFAULT 'excel'"),
                ("ai_model", "VARCHAR(50) DEFAULT ''"),
                ("ai_confidence", "DECIMAL(5,2) DEFAULT 0.00"),
            ]:
                try:
                    cursor.execute(f"ALTER TABLE `tb_import_log` ADD COLUMN `{col}` {col_type};")
                except Exception:
                    pass

            cursor.execute("""
                CREATE TABLE IF NOT EXISTS `tb_ai_cache` (
                    `id` INT AUTO_INCREMENT PRIMARY KEY,
                    `file_hash` VARCHAR(64) NOT NULL UNIQUE,
                    `file_type` VARCHAR(20) DEFAULT '',
                    `ai_model` VARCHAR(50) DEFAULT '',
                    `result_json` LONGTEXT DEFAULT NULL,
                    `created_at` TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
            """)

        conn.close()
        print("資料庫自愈巡檢完成：Gemini AI 多模態模組已 100% 準備就緒！")
    except Exception as e:
        print("資料庫巡檢通知:", e)

# ================= 3. 單號發號器 =================
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

# ================= 4. 全局設置與 Logo =================
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
            cursor.execute("UPDATE tb_layout_template SET logo_data = %s WHERE is_default = 1", (req.logo_data,))
        conn.close()
        return {"status": "success", "message": "公司 Logo 已成功永久保存！"}
    except Exception as e:
        if conn: conn.close()
        raise HTTPException(status_code=500, detail=str(e))

# ================= 5. 版式範本庫 API =================
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

# ================= 6. CRM 客戶管理 API =================
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

# ================= 7. 產品、匯率、報價單 API =================
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
        "USD_VND": 25967.0, "CNY_VND": 3874.0, "USD_CNY": 6.70,
        "USD_THB": 32.05, "USD_MYR": 3.9855
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

# ================= 8. 報價單儲存 =================
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

    conn = get_db()
    try:
        with conn.cursor() as cursor:
            final_total = Decimal(str(req.final_total_after_tax))
            terms_json_str = json.dumps(req.terms_list, ensure_ascii=False)

            cursor.execute("SELECT pair, rate FROM tb_exchange_rate")
            rate_map = {r['pair']: Decimal(str(r['rate'])) for r in cursor.fetchall()}
            rate_usd_vnd = rate_map.get("USD_VND", Decimal("25967.0"))
            rate_usd_cny = rate_map.get("USD_CNY", Decimal("6.70"))

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
                "status": "success", "message": msg, "quote_id": quote_id,
                "quote_no": final_quote_no, "final_total_after_tax": float(final_total)
            }
    except Exception as e:
        if conn: conn.close()
        print("資料庫保存異常詳情:", str(e))
        raise HTTPException(status_code=500, detail=f"資料庫保存出錯: {str(e)}")

# ================= 9. Excel 智能解析引擎 =================
class ExcelImportItem(BaseModel):
    product_code: str
    product_spec: str = ""
    unit: str = "个 / Cái"
    quantity: int = 1
    unit_price: float = 0.0
    note: str = ""
    confirmed: bool = True

class UnifiedImportRequest(BaseModel):
    file_name: str = "uploaded"
    import_mode: str = "library"
    import_source: str = "excel"
    ai_model: Optional[str] = ""
    ai_confidence: Optional[float] = 0.0
    customer_id: Optional[int] = 0
    customer_name: Optional[str] = ""
    items: List[ExcelImportItem] = []

def detect_header_row(rows):
    keywords = ['stt', '序號', '产品名称', '產品名稱', 'tên sản phẩm', 'product', 'model', '型號', 'đơn giá', '單價', 'unit price']
    for idx, row in enumerate(rows[:30]):
        row_text = ' '.join([str(cell).lower() for cell in row if cell is not None])
        match_count = sum(1 for kw in keywords if kw in row_text)
        if match_count >= 2:
            return idx
    return -1

def is_summary_row(row):
    row_text = ' '.join([str(cell).upper() for cell in row if cell is not None])
    summary_keywords = ['TỔNG', 'TOTAL', 'THUẾ', 'VAT', 'SUM', '稅前', '含稅', '增值稅', 'TỔNG GIÁ', 'TỔNG TRƯỚC', 'TỔNG SAU']
    return any(kw in row_text for kw in summary_keywords)

def is_valid_product_row(row):
    if not row or len(row) < 3:
        return False
    non_empty = [str(cell).strip() for cell in row if cell is not None and str(cell).strip() != '']
    if len(non_empty) < 3:
        return False
    if is_summary_row(row):
        return False
    has_number = False
    for cell in row:
        try:
            v = float(str(cell).replace(',', '').replace(' ', ''))
            if v > 0:
                has_number = True
                break
        except:
            continue
    return has_number

def smart_extract_products(rows):
    if not rows:
        return []
    header_idx = detect_header_row(rows)
    if header_idx < 0:
        header_idx = 0

    products = []
    for row in rows[header_idx + 1:]:
        if not row or not is_valid_product_row(row):
            continue
        cells = [str(c).strip() if c is not None else '' for c in row]

        product_code = cells[1] if len(cells) > 1 and cells[1] else ''
        product_spec = cells[3] if len(cells) > 3 and cells[3] else ''
        unit = cells[5] if len(cells) > 5 and cells[5] else '个 / Cái'
        quantity = 1
        for idx in [6, 5, 7]:
            if len(cells) > idx and cells[idx]:
                try:
                    quantity = int(float(cells[idx].replace(',', '')))
                    break
                except:
                    continue
        unit_price = 0.0
        for idx in [7, 6, 8]:
            if len(cells) > idx and cells[idx]:
                try:
                    unit_price = float(cells[idx].replace(',', '').replace(' ', ''))
                    break
                except:
                    continue
        note = cells[9] if len(cells) > 9 and cells[9] else ''

        if not product_code:
            for cell in cells[1:4]:
                if cell and not cell.replace('.', '').replace(',', '').isdigit():
                    product_code = cell
                    break

        if product_code and unit_price > 0:
            products.append({
                "product_code": product_code,
                "product_spec": product_spec,
                "unit": unit,
                "quantity": quantity,
                "unit_price": unit_price,
                "note": note,
                "confirmed": True
            })
    return products

@app.post("/api/excel/parse")
async def parse_excel_file(file: UploadFile = File(...)):
    try:
        content = await file.read()
        file_name = file.filename or "uploaded.xlsx"

        if file_name.lower().endswith('.csv'):
            df = pd.read_csv(io.BytesIO(content), header=None, dtype=str)
        else:
            xls = pd.ExcelFile(io.BytesIO(content))
            df = None
            for sheet_name in xls.sheet_names:
                temp_df = pd.read_excel(xls, sheet_name=sheet_name, header=None, dtype=str)
                if temp_df.shape[0] > 3:
                    df = temp_df
                    break
            if df is None:
                df = pd.read_excel(xls, sheet_name=0, header=None, dtype=str)

        rows = df.fillna('').values.tolist()
        products = smart_extract_products(rows)

        customer_name = ""
        for row in rows[:30]:
            row_text = ' '.join([str(c) for c in row if c])
            if 'khách hàng' in row_text.lower() or '客戶名稱' in row_text:
                for cell in reversed(row):
                    if cell and str(cell).strip() and 'khách' not in str(cell).lower() and '客戶' not in str(cell):
                        customer_name = str(cell).strip()
                        break
                if customer_name:
                    break

        return {
            "status": "success", "file_name": file_name, "total_rows": len(rows),
            "header_row_index": detect_header_row(rows), "detected_customer": customer_name,
            "products": products, "products_count": len(products)
        }
    except Exception as e:
        print("Excel 解析失敗:", str(e))
        raise HTTPException(status_code=500, detail=f"Excel 解析失敗: {str(e)}")

# ================= 10. Gemini AI 多模態解析（含自動重試） =================
AI_EXTRACTION_PROMPT = """你是專業的商業報價單數據提取助手。請從用戶提供的文件中，精確提取所有物料/產品明細。

**輸出要求**：
1. 只返回純 JSON，不要任何解釋文字
2. JSON 結構如下：
{
  "detected_customer": "客戶名稱（若有）",
  "detected_quote_no": "報價單號（若有）",
  "detected_date": "日期（若有）",
  "detected_currency": "VND/CNY/USD（推斷）",
  "confidence": 0.95,
  "products": [
    {
      "product_code": "產品型號（必填）",
      "product_spec": "規格/質量/品牌",
      "unit": "單位",
      "quantity": 1,
      "unit_price": 0.00,
      "note": "備註"
    }
  ]
}

**提取規則**：
- 表格中每行一個產品；型號為必填，沒有型號的行請跳過
- 單價只保留數字（去除 VND、,、空格等）
- 若同一型號有多個規格（如 ZH 3000v / 6000v），請分別列出
- 幣種判斷：VND 通常 4-7 位數、CNY 通常 2-4 位數、USD 通常 1-3 位數
- confidence 為你對整體識別準確度的評估（0-1）
- 若文件中找不到任何產品明細，products 返回空數組

**語言支持**：越南語、繁體中文、簡體中文、英文，請自動識別。
"""

def calculate_file_hash(file_bytes: bytes) -> str:
    return hashlib.sha256(file_bytes).hexdigest()

def check_ai_cache(file_hash: str) -> Optional[Dict]:
    try:
        conn = get_db()
        with conn.cursor() as cursor:
            cursor.execute("SELECT result_json FROM tb_ai_cache WHERE file_hash = %s LIMIT 1", (file_hash,))
            row = cursor.fetchone()
            conn.close()
            if row:
                return json.loads(row['result_json'])
    except Exception as e:
        print(f"快取查詢失敗：{e}")
    return None

def save_ai_cache(file_hash: str, file_type: str, result: Dict):
    try:
        conn = get_db()
        with conn.cursor() as cursor:
            cursor.execute("""
                INSERT INTO tb_ai_cache (file_hash, file_type, ai_model, result_json)
                VALUES (%s, %s, %s, %s)
                ON DUPLICATE KEY UPDATE result_json = VALUES(result_json)
            """, (file_hash, file_type, GEMINI_MODEL, json.dumps(result, ensure_ascii=False)))
        conn.close()
    except Exception as e:
        print(f"快取保存失敗：{e}")

async def gemini_extract_from_file(file_bytes: bytes, mime_type: str) -> Dict[str, Any]:
    """調用 Gemini 3.8 Flash 從文件提取物料信息（含 503/429 自動重試）"""
    if not GEMINI_AVAILABLE or not gemini_client:
        raise HTTPException(status_code=503, detail="Gemini AI 未配置，請設置 GEMINI_API_KEY 環境變數")

    contents = [
        AI_EXTRACTION_PROMPT,
        genai_types.Part.from_bytes(data=file_bytes, mime_type=mime_type)
    ]

    max_attempts = 8  # 最多 8 次嘗試
    last_error = None

    for attempt in range(max_attempts):
        try:
            print(f"[Gemini] 第 {attempt+1}/{max_attempts} 次嘗試（模型：{GEMINI_MODEL}）...")

            response = gemini_client.models.generate_content(
                model=GEMINI_MODEL,
                contents=contents,
                config=genai_types.GenerateContentConfig(
                    temperature=0.1,
                    max_output_tokens=8192,
                    response_mime_type="application/json",
                )
            )

            result_text = response.text.strip()
            result_text = re.sub(r'^```json\s*', '', result_text)
            result_text = re.sub(r'\s*```$', '', result_text)

            parsed = json.loads(result_text)
            parsed.setdefault("products", [])
            parsed.setdefault("detected_customer", "")
            parsed.setdefault("detected_quote_no", "")
            parsed.setdefault("detected_date", "")
            parsed.setdefault("detected_currency", "VND")
            parsed.setdefault("confidence", 0.8)

            cleaned_products = []
            for p in parsed.get("products", []):
                code = str(p.get("product_code", "")).strip()
                if not code:
                    continue
                cleaned_products.append({
                    "product_code": code,
                    "product_spec": str(p.get("product_spec", "")).strip(),
                    "unit": str(p.get("unit", "个 / Cái")).strip() or "个 / Cái",
                    "quantity": int(float(p.get("quantity", 1) or 1)),
                    "unit_price": float(str(p.get("unit_price", 0)).replace(",", "").replace(" ", "") or 0),
                    "note": str(p.get("note", "")).strip(),
                    "confirmed": True
                })
            parsed["products"] = cleaned_products
            print(f"[Gemini] ✅ 成功識別 {len(cleaned_products)} 項")
            return parsed

        except json.JSONDecodeError as e:
            raise HTTPException(status_code=500, detail=f"AI 返回格式異常：{str(e)}")

        except Exception as e:
            err_str = str(e)
            last_error = err_str

            # 503 / 429：等待後重試（指數退避）
            if "503" in err_str or "UNAVAILABLE" in err_str or "429" in err_str or "RESOURCE_EXHAUSTED" in err_str:
                if attempt < max_attempts - 1:
                    wait = (2 ** attempt) + _random.uniform(0, 1.5)
                    print(f"[Gemini] ⏳ 伺服器繁忙，等待 {wait:.1f}s 後重試...")
                    _time.sleep(wait)
                    continue
                else:
                    raise HTTPException(
                        status_code=503,
                        detail=f"Gemini 伺服器連續 {max_attempts} 次繁忙，請稍後 1-2 分鐘再試。（{err_str[:150]}）"
                    )

            # 404：模型不可用，立即拋出
            elif "404" in err_str or "NOT_FOUND" in err_str:
                raise HTTPException(
                    status_code=500,
                    detail=f"模型 {GEMINI_MODEL} 不可用，請確認您的 API Key 支援此模型。（{err_str[:150]}）"
                )

            # 編碼錯誤：Key 含有非 ASCII 字符
            elif isinstance(e, UnicodeEncodeError) or "codec can't encode" in err_str:
                raise HTTPException(
                    status_code=500,
                    detail="Gemini API Key 含有中文或全角等非 ASCII 字符，請重新設置純英文數字的 GEMINI_API_KEY 後重啟服務。"
                )

            # 其他錯誤
            else:
                raise HTTPException(status_code=500, detail=f"Gemini API 調用失敗：{err_str[:200]}")

    raise HTTPException(status_code=503, detail=f"Gemini 調用失敗：{last_error}")

@app.post("/api/ai/parse_file")
async def ai_parse_file(file: UploadFile = File(...)):
    try:
        content = await file.read()
        file_name = file.filename or "unknown"
        if len(content) > 20 * 1024 * 1024:
            raise HTTPException(status_code=400, detail="文件過大，請控制在 20MB 以內")

        mime_type, _ = mimetypes.guess_type(file_name)
        if not mime_type:
            if file_name.lower().endswith('.pdf'):
                mime_type = 'application/pdf'
            elif file_name.lower().endswith(('.jpg', '.jpeg')):
                mime_type = 'image/jpeg'
            elif file_name.lower().endswith('.png'):
                mime_type = 'image/png'
            elif file_name.lower().endswith('.webp'):
                mime_type = 'image/webp'
            else:
                raise HTTPException(status_code=400, detail=f"無法識別文件類型：{file_name}")

        file_hash = calculate_file_hash(content)
        cached = check_ai_cache(file_hash)
        if cached:
            cached["from_cache"] = True
            cached["file_name"] = file_name
            return {"status": "success", "data": cached}

        result = await gemini_extract_from_file(content, mime_type)
        result["from_cache"] = False
        result["file_name"] = file_name
        result["mime_type"] = mime_type
        result["file_size"] = len(content)
        save_ai_cache(file_hash, mime_type, result)
        return {"status": "success", "data": result}

    except HTTPException:
        raise
    except Exception as e:
        print(f"AI 解析失敗：{e}")
        raise HTTPException(status_code=500, detail=f"AI 解析失敗：{str(e)}")

@app.get("/api/ai/status")
def ai_status():
    return {
        "available": GEMINI_AVAILABLE and bool(GEMINI_API_KEY) and gemini_client is not None,
        "model": GEMINI_MODEL if GEMINI_AVAILABLE else None,
        "api_key_configured": bool(GEMINI_API_KEY),
        "supported_formats": ["jpg", "jpeg", "png", "webp", "gif", "pdf"] if GEMINI_AVAILABLE else []
    }

# ================= 11. 統一導入 API =================
@app.post("/api/import/unified")
def unified_import(req: UnifiedImportRequest):
    if not req.items:
        raise HTTPException(status_code=400, detail="沒有可導入的產品明細")

    conn = get_db()
    try:
        with conn.cursor() as cursor:
            cursor.execute("SELECT pair, rate FROM tb_exchange_rate")
            rate_map = {r['pair']: Decimal(str(r['rate'])) for r in cursor.fetchall()}
            rate_usd_vnd = rate_map.get("USD_VND", Decimal("25967.0"))

            new_count = 0
            existing_count = 0
            imported_products = []

            for item in req.items:
                if not item.confirmed:
                    continue
                clean_code = item.product_code.strip()
                clean_spec = item.product_spec.strip() if item.product_spec else ""
                if not clean_code:
                    continue

                cursor.execute(
                    "SELECT id, model, spec FROM tb_product WHERE model = %s AND spec = %s LIMIT 1",
                    (clean_code, clean_spec)
                )
                exist = cursor.fetchone()

                if exist:
                    existing_count += 1
                    imported_products.append({
                        "id": exist['id'], "model": exist['model'],
                        "spec": exist['spec'], "status": "existing"
                    })
                    continue

                prefix = "AI" if req.import_source == "ai" else "IMP"
                safe_tag = re.sub(r'[^\w\s-]', '', clean_code).replace(' ', '_')[:18]
                new_sku = f"{prefix}-{safe_tag}-{uuid.uuid4().hex[:8].upper()}"

                cursor.execute("""
                    INSERT INTO tb_product (
                        sku, brand, category, model, spec,
                        pcs_per_pallet, pcs_per_container, warranty_years, is_active
                    ) VALUES (%s, %s, %s, %s, %s, 1, 1, 1, 1)
                """, (
                    new_sku,
                    "AI導入" if req.import_source == "ai" else "Excel導入",
                    "通用工程輔材",
                    clean_code, clean_spec
                ))
                new_pid = cursor.lastrowid
                new_count += 1

                unit_price_vnd = Decimal(str(item.unit_price)) if item.unit_price > 0 else Decimal("0.00")
                base_usd = (unit_price_vnd / rate_usd_vnd).quantize(Decimal('0.01'), ROUND_HALF_UP) if rate_usd_vnd > 0 else Decimal("0.00")

                for reg in ['越南本地倉', '中國出港']:
                    for tier in ['大型經銷商', '中型經銷商', '中型安裝商', '小型安裝商']:
                        cursor.execute("""
                            INSERT INTO tb_price_matrix (
                                product_id, region, trade_term, customer_tier, currency, unit_price
                            ) VALUES (%s, %s, 'DAP', %s, 'USD', %s)
                            ON DUPLICATE KEY UPDATE unit_price=VALUES(unit_price)
                        """, (new_pid, reg, tier, float(base_usd)))

                imported_products.append({
                    "id": new_pid, "model": clean_code, "spec": clean_spec,
                    "unit_price_vnd": float(unit_price_vnd), "status": "new"
                })

            quotation_id = 0
            quotation_no = ""
            if req.import_mode == 'customer' and req.customer_id and req.customer_id > 0:
                now = datetime.datetime.now()
                prefix = f"BG-{now.strftime('%Y.%m.%d')}"
                cursor.execute("SELECT quote_no FROM tb_quotation WHERE quote_no LIKE %s", (f"{prefix}%",))
                all_existing = cursor.fetchall()
                max_seq = 0
                for row_it in all_existing:
                    m = re.search(r'-(\d+)$', row_it.get('quote_no', ''))
                    if m:
                        s = int(m.group(1))
                        if s > max_seq: max_seq = s
                quotation_no = f"{prefix}-{str(max_seq + 1).zfill(3)}"

                total_before_tax = sum(
                    (item.quantity * item.unit_price) for item in req.items if item.confirmed
                )
                vat_amount = total_before_tax * 0.1
                final_total = total_before_tax + vat_amount

                cursor.execute("""
                    INSERT INTO tb_quotation (
                        quote_no, customer_id, customer_name, customer_tier,
                        destination, quote_currency, markup_rate, tax_rate,
                        final_total_amount, valid_days, status
                    ) VALUES (%s, %s, %s, '中型經銷商', '越南本地倉 (DAP)', 'VND', 0.0, 10.0, %s, 3, '已生效')
                """, (quotation_no, req.customer_id, req.customer_name, float(final_total)))
                quotation_id = cursor.lastrowid

                for item in req.items:
                    if not item.confirmed:
                        continue
                    cursor.execute("""
                        INSERT INTO tb_quotation_item (
                            quotation_id, product_id, item_name, item_model, item_spec,
                            unit, quantity, applied_tier, base_price, quote_unit_price,
                            subtotal_amount, note
                        ) VALUES (%s, 0, %s, %s, %s, %s, %s, '中型經銷商', %s, %s, %s, %s)
                    """, (
                        quotation_id, item.product_code, item.product_code, item.product_spec,
                        item.unit, item.quantity, item.unit_price, item.unit_price,
                        item.quantity * item.unit_price, item.note
                    ))

            cursor.execute("""
                INSERT INTO tb_import_log (
                    file_name, import_mode, import_source, ai_model, ai_confidence,
                    customer_id, customer_name, total_products, new_products,
                    existing_products, quotation_id, raw_data_json
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """, (
                req.file_name, req.import_mode, req.import_source,
                req.ai_model or '', req.ai_confidence or 0.0,
                req.customer_id or 0, req.customer_name or '',
                len([i for i in req.items if i.confirmed]),
                new_count, existing_count, quotation_id,
                json.dumps([i.dict() for i in req.items], ensure_ascii=False)
            ))

            conn.close()
            return {
                "status": "success",
                "message": f"導入成功！新增 {new_count} 項，已存在 {existing_count} 項。",
                "new_count": new_count,
                "existing_count": existing_count,
                "quotation_id": quotation_id,
                "quotation_no": quotation_no,
                "imported_products": imported_products
            }
    except Exception as e:
        if conn: conn.close()
        print(f"導入失敗：{e}")
        raise HTTPException(status_code=500, detail=f"導入失敗：{str(e)}")

@app.get("/api/excel/import_history")
def get_import_history():
    try:
        conn = get_db()
        with conn.cursor() as cursor:
            cursor.execute("""
                SELECT id, file_name, import_mode, import_source, ai_model, ai_confidence,
                       customer_name, total_products, new_products, existing_products,
                       quotation_id, created_at
                FROM tb_import_log
                ORDER BY id DESC LIMIT 50
            """)
            rows = cursor.fetchall()
            for r in rows:
                if 'created_at' in r and r['created_at']:
                    r['created_at'] = str(r['created_at'])
            conn.close()
            return rows
    except Exception as e:
        print("獲取導入歷史失敗:", e)
        return []

# ================= 12. 前端主頁面 =================
@app.get("/", response_class=HTMLResponse)
def index():
    return """<!DOCTYPE html>
<html lang="zh-TW">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>GalaxyTeck 智慧快速報價系統 v4.0.2</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <script src="https://unpkg.com/vue@3/dist/vue.global.js"></script>
  <style>
    [v-cloak] { display: none !important; }

    /* ============ 設計變量：GalaxyTECK 品牌藏青 + 金，財務單據風格 ============ */
    :root {
      --navy: #1B2A4A; --navy-2: #2A3D66; --navy-3: #0F1A31;
      --gold: #C9A227; --gold-ink: #7A5F0B; --gold-wash: #FBF6E4;
      --ink: #172033; --muted: #5A6477; --faint: #8791A3;
      --rule: #D8DDE6; --rule-2: #B9C1CF; --ledger: #F1F3F7; --paper: #FFFFFF;
      --pos: #067647; --pos-wash: #E8F5EE; --neg: #B42318; --neg-wash: #FDECEA;
      --warn: #7A4B00; --warn-wash: #FFF4DB; --warn-line: #E0A93B;
      --font: "Microsoft YaHei UI", "Microsoft YaHei", "Segoe UI", "PingFang TC", "Noto Sans TC", system-ui, -apple-system, sans-serif;
    }

    .app-body { font-family: var(--font); background: var(--ledger); color: var(--ink); font-size: 13px; line-height: 1.45; -webkit-font-smoothing: antialiased; }
    .num { font-variant-numeric: tabular-nums lining-nums; font-feature-settings: "tnum" 1, "lnum" 1; }
    .wrap { max-width: 1360px; margin: 0 auto; padding: 0 20px; }
    .top-logo-img { max-height: 28px; width: auto; object-fit: contain; display: block; }
    .sheet-logo-img { max-height: 48px; width: auto; object-fit: contain; }
    .auto-expand-text { field-sizing: content; resize: none; }
    :focus-visible { outline: 2px solid var(--gold); outline-offset: 1px; }
    input[type=checkbox], input[type=radio] { accent-color: var(--navy); }

    /* ============ 按鈕 ============ */
    .btn { display: inline-flex; align-items: center; justify-content: center; gap: 6px; height: 32px; padding: 0 14px; border: 1px solid var(--rule-2); border-radius: 4px; background: #fff; color: var(--navy); font-size: 12px; font-weight: 600; cursor: pointer; white-space: nowrap; transition: background-color .12s, border-color .12s; }
    .btn:hover { background: var(--ledger); border-color: var(--faint); }
    .btn:disabled { opacity: .5; cursor: not-allowed; }
    .btn-primary { background: var(--navy); border-color: var(--navy); color: #fff; }
    .btn-primary:hover { background: var(--navy-2); border-color: var(--navy-2); }
    .btn-gold { background: var(--gold); border-color: var(--gold); color: var(--navy-3); }
    .btn-gold:hover { background: #D6B23A; border-color: #D6B23A; }
    .btn-danger { color: var(--neg); border-color: #E7B4AE; background: #fff; }
    .btn-danger:hover { background: var(--neg-wash); border-color: var(--neg); }
    .btn-onnav { background: transparent; border-color: rgba(255,255,255,.32); color: #fff; }
    .btn-onnav:hover { background: rgba(255,255,255,.12); border-color: rgba(255,255,255,.55); }
    .btn-sm { height: 26px; padding: 0 10px; font-size: 11px; }
    .link { color: var(--navy); font-size: 11px; font-weight: 600; text-decoration: underline; text-underline-offset: 2px; cursor: pointer; }
    .link:hover { color: var(--gold-ink); }
    .icon-btn { width: 28px; height: 28px; border-radius: 4px; color: inherit; font-size: 16px; line-height: 1; display: inline-flex; align-items: center; justify-content: center; cursor: pointer; }
    .icon-btn:hover { background: rgba(255,255,255,.14); }

    /* ============ 頂欄 ============ */
    .appbar { background: var(--navy); color: #fff; border-bottom: 3px solid var(--gold); }
    .appbar-top { display: flex; align-items: center; gap: 16px; padding: 12px 0 10px; flex-wrap: wrap; }
    .appbar-spacer { flex: 1; }
    .appbar-cmd { display: flex; gap: 8px; padding: 0 0 12px; flex-wrap: wrap; }
    .brand { cursor: pointer; display: block; }
    .brand-logo { display: flex; align-items: center; height: 38px; padding: 0 12px; background: #fff; border-radius: 4px; }
    .brand-fallback { display: flex; align-items: center; gap: 10px; height: 38px; padding: 0 12px; background: #fff; border-radius: 4px; color: var(--navy); font-size: 14px; }
    .brand-fallback i { font-style: normal; font-size: 11px; color: var(--muted); border-left: 1px solid var(--rule); padding-left: 10px; }
    .brand-title { display: flex; align-items: center; gap: 10px; font-size: 16px; font-weight: 700; }
    .brand-title .ver { font-size: 11px; font-weight: 600; color: var(--gold); border: 1px solid rgba(201,162,39,.6); border-radius: 3px; padding: 0 6px; }
    .brand-sub { font-size: 11px; color: rgba(255,255,255,.7); margin-top: 1px; }
    .seg { display: inline-flex; padding: 2px; border: 1px solid rgba(255,255,255,.22); border-radius: 4px; background: rgba(255,255,255,.07); }
    .seg-btn { height: 28px; padding: 0 12px; border-radius: 3px; font-size: 12px; font-weight: 600; color: rgba(255,255,255,.78); cursor: pointer; }
    .seg-btn:hover { color: #fff; }
    .seg-btn.is-on { background: #fff; color: var(--navy); }
    .seg-btn.is-on.is-warn { background: var(--gold); color: var(--navy-3); }

    /* ============ 財務版提示 ============ */
    .banner-warn { margin-top: 16px; padding: 10px 14px; background: var(--warn-wash); border: 1px solid var(--warn-line); border-left-width: 4px; border-radius: 4px; color: var(--warn); font-size: 12px; display: flex; align-items: center; gap: 12px; flex-wrap: wrap; }
    .banner-warn strong { font-weight: 700; }

    /* ============ 參數面板 ============ */
    .panel { background: #fff; border: 1px solid var(--rule); border-radius: 6px; margin-top: 16px; }
    .panel-head { display: flex; justify-content: space-between; align-items: baseline; gap: 12px; padding: 10px 16px; border-bottom: 1px solid var(--rule); }
    .panel-head h2 { font-size: 13px; font-weight: 700; color: var(--navy); }
    .panel-head .rate { font-size: 12px; color: var(--muted); }
    .panel-body { padding: 14px 16px 16px; display: flex; flex-direction: column; gap: 14px; }
    .grid-4 { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 12px 16px; }
    .grid-5 { display: grid; grid-template-columns: repeat(5, minmax(0, 1fr)); gap: 12px 16px; align-items: end; }
    .fld { position: relative; min-width: 0; }
    .fld-head { display: flex; justify-content: space-between; align-items: baseline; margin-bottom: 4px; }
    .lbl { display: block; font-size: 12px; color: var(--muted); margin-bottom: 4px; }
    .fld-head .lbl { margin-bottom: 0; }
    .inp { width: 100%; height: 32px; padding: 0 10px; border: 1px solid var(--rule-2); border-radius: 4px; background: #fff; color: var(--ink); font-size: 13px; }
    .inp:hover { border-color: var(--faint); }
    .inp:focus { outline: none; border-color: var(--gold-ink); box-shadow: 0 0 0 3px rgba(201,162,39,.28); }
    .inp-r { text-align: right; }
    .inp-key { font-weight: 700; color: var(--navy); background: var(--ledger); }
    .inp-strong { font-weight: 600; }
    .markup-box { background: var(--warn-wash); border: 1px solid var(--warn-line); border-radius: 4px; padding: 6px 8px 8px; }
    .markup-box .lbl { color: var(--warn); font-weight: 700; font-size: 11px; margin-bottom: 3px; }
    .markup-box .inp { height: 28px; border-color: var(--warn-line); font-weight: 700; }
    .toolrow { display: flex; flex-wrap: wrap; gap: 8px; align-items: center; padding-top: 14px; border-top: 1px solid var(--rule); }
    .toolrow .search { flex: 1; min-width: 260px; }
    .dropdown { position: absolute; left: 0; right: 0; top: 100%; margin-top: 4px; background: #fff; border: 1px solid var(--rule-2); border-radius: 4px; box-shadow: 0 10px 28px rgba(16,24,40,.16); max-height: 224px; overflow-y: auto; z-index: 40; }
    .dd-item { padding: 8px 10px; cursor: pointer; border-bottom: 1px solid var(--rule); }
    .dd-item:last-child { border-bottom: 0; }
    .dd-item:hover { background: var(--ledger); }
    .dd-name { font-size: 12px; font-weight: 700; color: var(--ink); }
    .dd-sub { font-size: 11px; color: var(--muted); }

    /* ============ 操作列（吸頂）：金額一直可見 ============ */
    .actionbar { position: sticky; top: 0; z-index: 30; margin-top: 16px; background: rgba(241,243,247,.95); backdrop-filter: blur(6px); border-bottom: 1px solid var(--rule); }
    .actionbar-inner { display: flex; align-items: center; gap: 24px; padding: 8px 0; flex-wrap: wrap; }
    .kv-mini { display: flex; flex-direction: column; min-width: 0; }
    .kv-mini .k { font-size: 11px; color: var(--muted); }
    .kv-mini .v { font-size: 14px; font-weight: 600; color: var(--ink); white-space: nowrap; }
    .kv-mini.docno .v { color: var(--navy); font-weight: 700; }
    .kv-mini.grand .v { font-size: 18px; font-weight: 700; color: var(--navy); border-bottom: 2px solid var(--gold); line-height: 1.25; }
    .kv-mini .v.pos { color: var(--pos); }
    .kv-mini .v.neg { color: var(--neg); }
    .totals-mini { display: flex; align-items: flex-end; gap: 24px; flex-wrap: wrap; }
    .actions { margin-left: auto; display: flex; gap: 8px; flex-wrap: wrap; }

    /* ============ A4 單據 ============ */
    .sheet-scroll { overflow-x: auto; padding: 16px 0 8px; }
    .a4-preview-card {
      width: 210mm; min-height: 297mm; margin: 0 auto; background: #fff;
      border: 1px solid var(--rule); border-top: 3px solid var(--navy); border-radius: 2px;
      padding: 12mm 14mm; box-sizing: border-box; position: relative;
      box-shadow: 0 1px 2px rgba(16,24,40,.06), 0 12px 32px rgba(16,24,40,.10);
    }
    .a4-preview-card input, .a4-preview-card textarea { background: transparent; border: 0; border-radius: 2px; color: inherit; font: inherit; padding: 1px 3px; outline: none; min-width: 0; }
    .a4-preview-card input:hover, .a4-preview-card textarea:hover { background: rgba(27,42,74,.05); }
    .a4-preview-card input:focus, .a4-preview-card textarea:focus { background: var(--gold-wash); box-shadow: inset 0 -2px 0 var(--gold); }
    .a4-preview-card input[type=number] { -moz-appearance: textfield; appearance: textfield; }
    .a4-preview-card input[type=number]::-webkit-inner-spin-button, .a4-preview-card input[type=number]::-webkit-outer-spin-button { -webkit-appearance: none; margin: 0; }

    .header-box { border-bottom: 1.5px solid var(--navy); padding-bottom: 10px; margin-bottom: 12px; }
    .hb-row { display: flex; justify-content: space-between; align-items: flex-start; gap: 16px; }
    .hb-logo { max-width: 240px; }
    .logo-g { font-size: 26px; font-weight: 800; color: var(--gold); margin-right: 2px; }
    .logo-word { font-size: 20px; font-weight: 800; color: var(--navy); }
    .hb-names { flex: 1; text-align: right; }
    .hb-names textarea { width: 100%; text-align: right; display: block; }
    .name-vi { font-size: 13px; font-weight: 700; color: var(--navy); }
    .name-cn { font-size: 11px; color: var(--muted); margin-top: 2px; }
    .co-meta { margin-top: 8px; padding-top: 6px; border-top: 1px solid var(--rule); display: grid; grid-template-columns: 1fr 1fr; gap: 2px 16px; font-size: 10px; color: var(--muted); }
    .co-meta > div { display: flex; align-items: baseline; gap: 4px; min-width: 0; }
    .co-meta input { flex: 1; width: 100%; color: var(--ink); }
    .co-meta .span-2 { grid-column: span 2; }
    .co-meta .ta-r input { text-align: right; }

    .doc-title-wrap { text-align: center; margin: 14px 0 12px; }
    .doc-title { width: 100%; text-align: center; font-size: 20px; font-weight: 800; color: var(--navy); letter-spacing: .02em; }
    .badge-secret { display: inline-block; margin-bottom: 4px; padding: 1px 10px; background: var(--warn-wash); border: 1px solid var(--warn-line); border-radius: 3px; color: var(--warn); font-size: 10px; font-weight: 700; }

    .info-grid { display: grid; grid-template-columns: 1fr 1fr; border: 1px solid var(--rule-2); margin-bottom: 12px; font-size: 11px; }
    .info-grid > div { padding: 8px 10px; }
    .info-grid > div + div { border-left: 1px solid var(--rule-2); }
    .info-h { font-size: 11px; font-weight: 700; color: var(--navy); border-bottom: 1px solid var(--rule); padding-bottom: 4px; margin-bottom: 6px; }
    .info-row { display: flex; align-items: baseline; gap: 8px; padding: 2px 0; }
    .info-row .k { width: 64px; flex: none; color: var(--muted); }
    .info-row.between { justify-content: space-between; }
    .info-row.between .k { width: auto; }
    .info-row input { flex: 1; width: 100%; }
    .info-row.between input { flex: none; width: auto; text-align: right; }
    .due { color: var(--neg); font-weight: 600; }

    .fin-box { margin-bottom: 12px; padding: 10px 12px; background: var(--ledger); border: 1px solid var(--rule-2); border-left: 3px solid var(--warn-line); }
    .fin-box h3 { font-size: 11px; font-weight: 700; color: var(--navy); border-bottom: 1px solid var(--rule); padding-bottom: 4px; margin-bottom: 8px; }
    .fin-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 8px; }
    .fin-cell { background: #fff; border: 1px solid var(--rule); padding: 6px 8px; text-align: right; }
    .fin-cell .k { font-size: 10px; color: var(--muted); text-align: left; }
    .fin-cell .v { font-size: 12px; font-weight: 700; }
    .pos { color: var(--pos); }
    .neg { color: var(--neg); }

    .greeting { font-size: 10px; font-style: italic; color: var(--muted); margin-bottom: 6px; }
    .greeting input { width: 100%; font-style: italic; }

    .item-table { width: 100%; border-collapse: collapse; font-size: 11px; margin-bottom: 14px; border-top: 2px solid var(--navy); border-bottom: 1px solid var(--navy); }
    .item-table thead th { background: var(--ledger); color: var(--navy); font-weight: 700; padding: 6px 6px; border-bottom: 1px solid var(--navy); text-align: left; vertical-align: bottom; line-height: 1.25; }
    .item-table th.c, .item-table td.c { text-align: center; }
    .item-table th.r, .item-table td.r { text-align: right; }
    .item-table tbody td { padding: 5px 6px; border-bottom: 1px solid var(--rule); vertical-align: middle; }
    .item-table tbody tr:last-child td { border-bottom: 0; }
    .item-table tbody tr:hover td { background: #FAFBFD; }
    .item-table td input { width: 100%; text-overflow: ellipsis; }
    .item-table td.r input, .item-table td.c input.qty { text-align: right; }
    .item-table .col-markup { background: var(--warn-wash); }
    .item-table thead th.col-markup { background: #FCE9BE; }
    .item-table .col-cost { background: var(--ledger); }
    .item-table .col-total { font-weight: 700; }
    .item-table .col-profit { color: var(--pos); font-weight: 600; }
    .item-table .markup-in { width: 52px; text-align: center; border: 1px solid var(--warn-line); background: #fff; }
    .item-table .rm { color: var(--neg); font-weight: 700; cursor: pointer; width: 22px; height: 22px; border-radius: 3px; }
    .item-table .rm:hover { background: var(--neg-wash); }
    .empty-row td { text-align: center; padding: 28px 8px !important; color: var(--faint); }

    .totals-wrap { display: flex; justify-content: flex-end; margin-bottom: 14px; }
    .totals { width: 300px; font-size: 11.5px; }
    .totals .row { display: flex; justify-content: space-between; gap: 12px; padding: 5px 10px; border-bottom: 1px solid var(--rule); }
    .totals .row.grand { margin-top: 2px; border-top: 1.5px solid var(--navy); border-bottom: 3px double var(--navy); background: var(--gold-wash); color: var(--navy); font-weight: 700; font-size: 13px; padding: 7px 10px; }

    .terms-box { margin-bottom: 14px; padding: 10px 12px; background: #FAFBFD; border: 1px solid var(--rule); font-size: 10px; }
    .terms-head { display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid var(--rule); padding-bottom: 4px; margin-bottom: 6px; }
    .terms-head .t { font-weight: 700; color: var(--navy); font-size: 11px; }
    .term-row { display: flex; align-items: flex-start; gap: 4px; }
    .term-row .n { padding-top: 2px; font-weight: 700; color: var(--navy); }
    .term-row input { flex: 1; font-size: 10px; }
    .term-ops { display: flex; gap: 2px; opacity: 0; }
    .term-row:hover .term-ops, .term-row:focus-within .term-ops { opacity: 1; }
    .term-ops button { color: var(--faint); padding: 0 4px; cursor: pointer; }
    .term-ops button:hover { color: var(--navy); }
    .term-ops button:disabled { opacity: .3; cursor: default; }
    .term-ops .del { color: var(--neg); font-weight: 700; }

    .signature-box { display: grid; grid-template-columns: 1fr 1fr; gap: 24px; text-align: center; font-size: 11px; padding-top: 6px; border-top: 1px solid var(--navy); }
    .signature-box .who { font-weight: 700; color: var(--navy); margin-bottom: 4px; }
    .sig-space { height: 56px; display: flex; align-items: flex-end; justify-content: center; color: var(--faint); font-style: italic; font-size: 9px; border-bottom: 1px solid var(--rule-2); }
    .sig-seal { height: 56px; display: flex; flex-direction: column; justify-content: flex-end; align-items: center; font-size: 10px; border-bottom: 1px solid var(--rule-2); }
    .sig-seal .co { font-weight: 700; }
    .sig-seal .dir { font-weight: 600; }
    .sheet-foot { margin-top: 12px; padding-top: 6px; border-top: 1px solid var(--rule); display: flex; justify-content: space-between; font-size: 8px; color: var(--faint); }

    /* ============ 彈窗 ============ */
    .modal-mask { position: fixed; inset: 0; background: rgba(15,26,49,.55); display: flex; align-items: center; justify-content: center; padding: 16px; z-index: 60; }
    .modal { background: #fff; border-radius: 6px; box-shadow: 0 24px 64px rgba(15,26,49,.38); width: 100%; max-height: 90vh; display: flex; flex-direction: column; overflow: hidden; }
    .modal-xl { max-width: 1024px; }
    .modal-lg { max-width: 896px; }
    .modal-head { display: flex; justify-content: space-between; align-items: center; gap: 12px; padding: 10px 12px 10px 20px; background: var(--navy); color: #fff; border-bottom: 3px solid var(--gold); }
    .modal-head .ttl { font-size: 14px; font-weight: 700; }
    .modal-head .sub { font-size: 12px; font-weight: 600; margin-left: 10px; }
    .modal-head .sub.ok { color: #7FD6A6; }
    .modal-head .sub.bad { color: #F4C46A; }
    .modal-body { padding: 16px 20px; display: flex; flex-direction: column; gap: 12px; overflow: hidden; flex: 1; min-height: 0; }
    .modal-foot { display: flex; justify-content: flex-end; gap: 8px; padding-top: 12px; border-top: 1px solid var(--rule); }
    .note { padding: 10px 12px; border: 1px solid var(--rule); border-radius: 4px; background: var(--ledger); font-size: 12px; }
    .note-warn { background: var(--warn-wash); border-color: var(--warn-line); color: var(--warn); }
    .note-ok { background: var(--pos-wash); border-color: #A8D8BE; color: var(--pos); font-weight: 600; }
    .note-info { background: #EEF2FA; border-color: #C4CFE6; }
    .note-title { font-weight: 700; color: var(--navy); margin-bottom: 6px; font-size: 12px; }
    .note-warn .note-title { color: var(--warn); }
    .note code { display: block; margin-top: 6px; padding: 6px 8px; background: #fff; border: 1px solid var(--rule); border-radius: 3px; font-size: 11px; color: var(--ink); overflow-x: auto; }
    .radio-row { display: flex; gap: 20px; font-size: 12px; flex-wrap: wrap; }
    .radio-row label { display: flex; align-items: center; gap: 6px; cursor: pointer; }
    .dropzone { padding: 16px; border: 1.5px dashed var(--rule-2); border-radius: 4px; background: var(--ledger); text-align: center; font-size: 12px; color: var(--muted); }
    .dropzone .fname { font-weight: 700; color: var(--pos); font-size: 12px; }
    .dropzone .hint { font-size: 11px; color: var(--muted); margin-top: 8px; }
    .dz-actions { display: flex; justify-content: center; gap: 8px; margin-top: 10px; }
    .result-wrap { display: flex; flex-direction: column; gap: 8px; flex: 1; min-height: 0; }
    .result-bar { display: flex; justify-content: space-between; align-items: center; gap: 12px; padding: 8px 12px; background: var(--pos-wash); border: 1px solid #A8D8BE; border-radius: 4px; font-size: 12px; }
    .result-bar .stats { display: flex; gap: 16px; flex-wrap: wrap; font-weight: 600; color: var(--pos); }
    .result-bar .stats .m { color: var(--navy); }
    .tbl-wrap { border: 1px solid var(--rule); border-radius: 4px; overflow: auto; flex: 1; min-height: 120px; }
    .tbl { width: 100%; border-collapse: collapse; font-size: 12px; }
    .tbl thead th { position: sticky; top: 0; z-index: 1; background: var(--ledger); color: var(--navy); font-weight: 700; padding: 8px 10px; border-bottom: 1px solid var(--navy); text-align: left; white-space: nowrap; }
    .tbl td { padding: 7px 10px; border-bottom: 1px solid var(--rule); vertical-align: middle; }
    .tbl tbody tr:hover { background: #FAFBFD; }
    .tbl .r { text-align: right; }
    .tbl .c { text-align: center; }
    .tbl .dim { color: var(--muted); }
    .tbl .key { font-weight: 700; color: var(--navy); }
    .tbl .row-off { background: var(--ledger); opacity: .55; }
    .cell-inp { width: 100%; height: 28px; padding: 0 6px; border: 1px solid var(--rule); border-radius: 3px; background: #fff; font-size: 12px; }
    .cell-inp:focus { outline: none; border-color: var(--gold-ink); box-shadow: 0 0 0 2px rgba(201,162,39,.28); }
    .cell-inp.sm { width: 64px; }
    .cell-inp.md { width: 96px; }
    .cell-inp.tc { text-align: center; }
    .cell-inp.tr { text-align: right; }
    .chip { display: inline-block; padding: 1px 8px; border-radius: 10px; font-size: 11px; font-weight: 600; background: var(--ledger); color: var(--navy); border: 1px solid var(--rule); }
    .chip-pos { background: var(--pos-wash); color: var(--pos); border-color: #A8D8BE; }
    .cust-form { padding: 14px; background: var(--ledger); border: 1px solid var(--rule); border-radius: 4px; display: flex; flex-direction: column; gap: 10px; font-size: 12px; }
    .cust-form .hd { display: flex; justify-content: space-between; align-items: center; font-weight: 700; color: var(--navy); }
    .cust-grid { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 10px; }
    .cust-grid .span-2 { grid-column: span 2; }
    .cust-orders { padding: 12px; background: #EEF2FA; border: 1px solid #C4CFE6; border-radius: 4px; display: flex; flex-direction: column; gap: 8px; }
    .cust-orders .hd { display: flex; justify-content: space-between; align-items: center; font-weight: 700; font-size: 13px; color: var(--navy); }
    .cust-orders .tbl-wrap { background: #fff; max-height: 224px; flex: none; }
    .row-actions { display: flex; justify-content: center; gap: 4px; }
    .empty-cell { text-align: center; padding: 28px 10px !important; color: var(--faint); }

    @media screen {
      .item-table .col-name { min-width: 120px; }
      .item-table .col-spec { min-width: 130px; }
      .a4-preview-card[data-view="finance"] { width: min(100%, 1180px); }
    }

    @media (max-width: 960px) {
      .grid-4, .grid-5 { grid-template-columns: repeat(2, minmax(0, 1fr)); }
      .cust-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
      .actions { margin-left: 0; }
    }
    @media (max-width: 560px) {
      .grid-4, .grid-5, .cust-grid { grid-template-columns: 1fr; }
      .cust-grid .span-2 { grid-column: auto; }
    }
    @media (prefers-reduced-motion: reduce) { * { transition: none !important; } }

    /* ============ 列印（A4 防跑版規則完整保留） ============ */
    @page { size: A4 portrait; margin: 12mm 14mm; }
    @media print {
      .no-print, header, nav { display: none !important; }
      body { background: white !important; font-size: 10pt !important; margin: 0 !important; padding: 0 !important; }
      .sheet-scroll { overflow: visible !important; padding: 0 !important; }
      .print-sheet { width: 100% !important; max-width: 100% !important; min-height: 0 !important; margin: 0 !important; padding: 0 !important; border: none !important; box-shadow: none !important; background: white !important; }
      input, textarea, select { border: none !important; background: transparent !important; padding: 0 !important; box-shadow: none !important; }
      table { page-break-inside: auto !important; width: 100% !important; border-collapse: collapse !important; }
      thead { display: table-header-group !important; }
      tr { page-break-inside: avoid !important; }
      .keep-together { page-break-inside: avoid !important; }
      .compact-mode .header-box { margin-bottom: 2mm !important; }
      .compact-mode .item-table th, .compact-mode .item-table td { padding-top: 2px !important; padding-bottom: 2px !important; }
      .item-table thead th, .totals .row.grand, .fin-box, .terms-box, .info-grid { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    }
  </style>
</head>
<body class="app-body min-h-screen">
<div id="app" v-cloak>

  <header class="appbar no-print">
    <div class="wrap">
      <div class="appbar-top">
        <label class="brand" title="點擊更換 Logo">
          <span v-if="companyInfo.logo_data" class="brand-logo"><img :src="companyInfo.logo_data" class="top-logo-img"></span>
          <span v-else class="brand-fallback"><b>GalaxyTECK</b><i>Change</i></span>
          <input type="file" @change="handleLogoUpload" accept="image/*" class="hidden">
        </label>
        <div>
          <div class="brand-title">
            <span>{{ t('workspace') }}</span>
            <span class="ver num">v4.0.2 AI</span>
          </div>
          <div class="brand-sub">GALAXY VIỆT NAM | MST: 2301296587</div>
        </div>
        <div class="appbar-spacer"></div>
        <div class="seg">
          <button @click="setLanguage('vi')" :class="['seg-btn', currentLang==='vi' ? 'is-on' : '']">Vi</button>
          <button @click="setLanguage('zh')" :class="['seg-btn', currentLang==='zh' ? 'is-on' : '']">中</button>
          <button @click="setLanguage('en')" :class="['seg-btn', currentLang==='en' ? 'is-on' : '']">En</button>
        </div>
        <div class="seg">
          <button @click="viewMode='customer'" :class="['seg-btn', viewMode==='customer' ? 'is-on' : '']">{{ t('customer_view') }}</button>
          <button @click="viewMode='finance'" :class="['seg-btn', viewMode==='finance' ? 'is-on is-warn' : '']">{{ t('finance_view') }}</button>
        </div>
      </div>
      <div class="appbar-cmd">
        <button @click="openAiImportModal" class="btn btn-onnav">AI 智能導入</button>
        <button @click="openExcelImportModal" class="btn btn-onnav">Excel 導入</button>
        <button @click="openCustomerModal" class="btn btn-onnav">CRM</button>
        <button @click="openHistoryModal" class="btn btn-onnav">歷史 ({{ historyList.length }})</button>
      </div>
    </div>
  </header>

  <div class="wrap">

    <div v-if="viewMode==='finance'" class="no-print banner-warn">
      <span><strong>【內部財務審核版】</strong> 顯示成本、調價率、利潤。嚴禁發送給客戶！</span>
      <button @click="viewMode='customer'" class="link">切換回客戶版</button>
    </div>

    <section class="no-print panel">
      <div class="panel-head">
        <h2>商業計價與交付參數</h2>
        <span class="rate num">匯率：1 USD = {{ rates['USD_VND'] || 25967 }} VND</span>
      </div>
      <div class="panel-body">
        <div class="grid-4">
          <div class="fld">
            <div class="fld-head">
              <label class="lbl">報價單號</label>
              <button @click="fetchNextQuoteNo" class="link">換新單號</button>
            </div>
            <input v-model="form.quote_no" class="inp inp-key num">
          </div>
          <div class="fld relative">
            <div class="fld-head">
              <label class="lbl">客戶名稱</label>
              <button @click="openCustomerModal" class="link">CRM選客戶</button>
            </div>
            <input v-model="form.customer_name" @focus="showCustomerDropdown=true" placeholder="輸入或選擇客戶..." class="inp inp-strong">
            <div v-if="showCustomerDropdown && matchedCustomers.length>0" class="dropdown">
              <div v-for="c in matchedCustomers" :key="c.id" @click="selectCustomer(c)" class="dd-item">
                <div class="dd-name">{{ c.company_name }}</div>
                <div class="dd-sub">{{ c.contact_person }} · {{ c.phone }}</div>
              </div>
            </div>
          </div>
          <div class="fld">
            <label class="lbl">聯絡人</label>
            <input v-model="form.contact_person" class="inp">
          </div>
          <div class="fld">
            <label class="lbl">電話/郵箱</label>
            <input v-model="form.phone_email" class="inp">
          </div>
        </div>

        <div class="grid-5">
          <div class="fld">
            <label class="lbl">交付場景</label>
            <select v-model="form.delivery_scenario" class="inp">
              <option value="越南本地倉">越南本地倉 (DAP)</option>
              <option value="中國出港">中國出港 (FOB)</option>
            </select>
          </div>
          <div class="fld">
            <label class="lbl">價格階梯</label>
            <select v-model="form.customer_tier" class="inp">
              <option>大型經銷商</option><option>中型經銷商</option>
              <option>中型安裝商</option><option>小型安裝商</option>
            </select>
          </div>
          <div class="fld">
            <label class="lbl">結算幣種</label>
            <select v-model="form.target_currency" @change="onCurrencyChange" class="inp inp-key">
              <option value="VND">VND</option><option value="CNY">CNY</option><option value="USD">USD</option>
            </select>
          </div>
          <div class="fld">
            <label class="lbl">增值稅 VAT (%)</label>
            <input type="number" v-model.number="form.vat_rate_pct" class="inp inp-r num">
          </div>
          <div class="markup-box">
            <label class="lbl">全局調價 (%)</label>
            <input type="number" step="0.5" v-model.number="form.markup_pct" class="inp inp-r num">
          </div>
        </div>

        <div class="toolrow">
          <button @click="openSearchModal" class="btn btn-primary">產品庫</button>
          <input v-model="quickSearchKeyword" @keyup.enter="quickAddFirst" placeholder="輸入型號搜尋 (Enter 快速選入)..." class="inp search">
          <button @click="resetCustomMarkups" v-if="hasAnyCustomMarkup" class="btn">恢復全局調價</button>
          <button @click="addCustomRow" class="btn">+ 自定義物料</button>
        </div>
      </div>
    </section>
  </div>

  <div class="no-print actionbar">
    <div class="wrap">
      <div class="actionbar-inner">
        <div class="kv-mini docno">
          <span class="k">單號</span>
          <span class="v num">{{ form.quote_no || '未生成' }}</span>
        </div>
        <div class="totals-mini">
          <div class="kv-mini">
            <span class="k">稅前總額</span>
            <span class="v num">{{ formatCurrency(totalBeforeTax) }}</span>
          </div>
          <div class="kv-mini">
            <span class="k">VAT {{ form.vat_rate_pct }}%</span>
            <span class="v num">{{ formatCurrency(vatAmount) }}</span>
          </div>
          <div class="kv-mini grand">
            <span class="k">含稅總額 ({{ form.target_currency }})</span>
            <span class="v num">{{ formatCurrency(finalTotalAfterTax) }}</span>
          </div>
          <template v-if="viewMode==='finance'">
            <div class="kv-mini">
              <span class="k">毛利額</span>
              <span :class="['v num', totalGrossProfit>=0 ? 'pos' : 'neg']">{{ formatCurrency(totalGrossProfit) }}</span>
            </div>
            <div class="kv-mini">
              <span class="k">毛利率</span>
              <span class="v num">{{ grossMarginPercent.toFixed(2) }}%</span>
            </div>
          </template>
        </div>
        <div class="actions">
          <button @click="handleNewQuote" class="btn">新建空白單</button>
          <button v-if="currentQuoteId" @click="saveToServer(true)" class="btn">另存為新單</button>
          <button @click="saveToServer(false)" class="btn btn-gold">{{ currentQuoteId ? '覆蓋更新' : '保存入庫' }}</button>
          <button @click="printOffer" class="btn btn-primary">列印</button>
        </div>
      </div>
    </div>
  </div>

  <div class="sheet-scroll">
  <div :data-view="viewMode" :class="['a4-preview-card print-sheet', printDensity === 'ultra' || tableRows.length <= 6 ? 'compact-mode' : '']">
    <div class="header-box">
      <div class="hb-row">
        <div class="hb-logo">
          <div v-if="companyInfo.logo_data"><img :src="companyInfo.logo_data" class="sheet-logo-img"></div>
          <div v-else>
            <span class="logo-g">G</span><span class="logo-word">GalaxyTECK</span>
          </div>
        </div>
        <div class="hb-names">
          <textarea v-model="companyInfo.name_vi" rows="1" class="auto-expand-text name-vi"></textarea>
          <textarea v-model="companyInfo.name_cn" rows="1" class="auto-expand-text name-cn"></textarea>
        </div>
      </div>
      <div class="co-meta">
        <div>MST: <input v-model="companyInfo.tax_code" class="num"></div>
        <div class="ta-r"><input v-model="companyInfo.contact_info"></div>
        <div class="span-2">Địa chỉ: <input v-model="companyInfo.address"></div>
        <div class="span-2"><input v-model="companyInfo.bank_info" class="num"></div>
      </div>
    </div>

    <div class="doc-title-wrap">
      <div v-if="viewMode==='finance'" class="badge-secret">內部機密</div>
      <input v-model="layoutConfig.title_text" class="doc-title">
    </div>

    <div class="info-grid">
      <div>
        <div class="info-h">客戶信息</div>
        <div class="info-row"><span class="k">客戶：</span><input v-model="form.customer_name" class="inp-strong" style="font-weight:600"></div>
        <div class="info-row"><span class="k">聯絡人：</span><input v-model="form.contact_person"></div>
        <div class="info-row"><span class="k">電話：</span><input v-model="form.phone_email"></div>
      </div>
      <div>
        <div class="info-h">報價信息</div>
        <div class="info-row between"><span class="k">單號：</span><input v-model="form.quote_no" class="num" style="font-weight:700"></div>
        <div class="info-row between"><span class="k">日期：</span><span class="num">{{ todayDate }}</span></div>
        <div class="info-row between"><span class="k">報價人：</span><input v-model="form.sales_rep"></div>
        <div class="info-row between"><span class="k">有效期：</span><span class="due num">{{ validUntilDate }} ({{ form.valid_days }} 天)</span></div>
      </div>
    </div>

    <div v-if="viewMode==='finance'" class="fin-box keep-together">
      <h3>內部利潤分析</h3>
      <div class="fin-grid num">
        <div class="fin-cell"><div class="k">總成本</div><div class="v">{{ formatCurrency(totalBaseCost) }}</div></div>
        <div class="fin-cell"><div class="k">稅前營收</div><div class="v" style="color:var(--navy)">{{ formatCurrency(totalBeforeTax) }}</div></div>
        <div class="fin-cell"><div class="k">毛利額</div><div :class="['v', totalGrossProfit>=0?'pos':'neg']">{{ formatCurrency(totalGrossProfit) }}</div></div>
        <div class="fin-cell"><div class="k">毛利率</div><div class="v pos">{{ grossMarginPercent.toFixed(2) }}%</div></div>
      </div>
    </div>

    <div class="greeting">
      <input v-model="layoutConfig.greeting_text">
    </div>

    <table class="item-table">
      <thead>
        <tr>
          <th class="c" style="width:32px">STT<br>序號</th>
          <th class="col-name">產品名稱</th>
          <th class="col-spec">規格型號</th>
          <th class="c" style="width:84px">單位</th>
          <th class="c" style="width:44px">數量</th>
          <th v-if="viewMode==='finance'" class="r col-cost" style="width:80px">底價成本</th>
          <th class="no-print c col-markup" style="width:60px">獨立調價</th>
          <th class="r" style="width:92px">對外單價</th>
          <th class="r" style="width:108px">總價</th>
          <th v-if="viewMode==='finance'" class="r" style="width:80px">毛利</th>
          <th class="no-print c" style="width:30px">操作</th>
        </tr>
      </thead>
      <tbody>
        <tr v-for="(row, idx) in tableRows" :key="idx">
          <td class="c num">{{ idx + 1 }}</td>
          <td><input v-model="row.name" style="font-weight:600"></td>
          <td><input v-model="row.spec" style="font-size:10px;color:var(--muted)"></td>
          <td class="c"><input v-model="row.unit" style="text-align:center"></td>
          <td class="c"><input type="number" v-model.number="row.quantity" class="qty num" style="width:36px"></td>
          <td v-if="viewMode==='finance'" class="r col-cost"><input type="number" v-model.number="row.base_price" class="num" style="width:70px"></td>
          <td class="no-print c col-markup">
            <input type="number" step="0.5" v-model.number="row.custom_markup" :placeholder="form.markup_pct" class="markup-in num">
          </td>
          <td class="r">
            <input type="number" :value="calcRowFinalUnitPrice(row)" @change="onManualUnitPriceChange(row, $event.target.value)" class="num" style="width:82px;font-weight:600">
          </td>
          <td class="r num col-total">{{ formatCurrency(calcRowSubtotal(row)) }}</td>
          <td v-if="viewMode==='finance'" class="r num col-profit">{{ formatCurrency(calcRowProfit(row)) }}</td>
          <td class="no-print c"><button @click="removeRow(idx)" class="rm" title="刪除此行">✕</button></td>
        </tr>
        <tr v-if="tableRows.length===0" class="empty-row">
          <td :colspan="viewMode==='finance' ? 11 : 9">清單為空，請新增產品</td>
        </tr>
      </tbody>
    </table>

    <div class="keep-together totals-wrap">
      <div class="totals num">
        <div class="row">
          <span>稅前總額：</span>
          <span style="font-weight:600">{{ formatCurrency(totalBeforeTax) }} {{ form.target_currency }}</span>
        </div>
        <div class="row">
          <span>VAT ({{ form.vat_rate_pct }}%)：</span>
          <span>{{ formatCurrency(vatAmount) }} {{ form.target_currency }}</span>
        </div>
        <div class="row grand">
          <span>含稅總額：</span>
          <span>{{ formatCurrency(finalTotalAfterTax) }} {{ form.target_currency }}</span>
        </div>
      </div>
    </div>

    <div class="terms-box keep-together">
      <div class="terms-head">
        <div class="t">商務條款及備註：</div>
        <button @click="addTermRow" class="no-print btn btn-sm btn-primary">+ 新增</button>
      </div>
      <div>
        <div v-for="(term, tIdx) in termsList" :key="tIdx" class="term-row">
          <span class="n num">{{ tIdx + 1 }})</span>
          <input v-model="termsList[tIdx]">
          <div class="no-print term-ops">
            <button @click="moveTermUp(tIdx)" :disabled="tIdx===0" title="上移">▲</button>
            <button @click="moveTermDown(tIdx)" :disabled="tIdx===termsList.length-1" title="下移">▼</button>
            <button @click="removeTermRow(tIdx)" class="del" title="刪除">✕</button>
          </div>
        </div>
      </div>
    </div>

    <div class="signature-box keep-together">
      <div>
        <div class="who">{{ viewMode==='customer' ? '客戶確認簽字蓋章' : '製單人' }}</div>
        <div class="sig-space">(簽字蓋章)</div>
      </div>
      <div>
        <div class="who">{{ viewMode==='customer' ? '賣方確認蓋章' : '總經理審批' }}</div>
        <div class="sig-seal">
          <div class="co">{{ companyInfo.name_vi }}</div>
          <div class="dir">GIÁM ĐỐC: WANG XIAN LI</div>
        </div>
      </div>
    </div>

    <div class="keep-together sheet-foot num">
      <span>GalaxyTeck v4.0.2 · Gemini 3.8 Flash · Page 1/1</span>
      <span>Copyright © 2024-2026 GALAXY VIETNAM CO., LTD.</span>
    </div>
  </div>
  </div>

  <!-- ============ AI 智能導入 ============ -->
  <div v-show="showAiModal" class="no-print modal-mask">
    <div class="modal modal-xl">
      <div class="modal-head">
        <div style="display:flex;align-items:center">
          <span class="ttl">AI 智能導入（Gemini 3.8 Flash）</span>
          <span v-if="aiStatus.available" class="sub ok">已啟用 · {{ aiStatus.model }}</span>
          <span v-else class="sub bad">未配置</span>
        </div>
        <button @click="closeAiModal" class="icon-btn" title="關閉">✕</button>
      </div>
      <div class="modal-body">
        <div v-if="!aiStatus.available" class="note note-warn">
          <div class="note-title">AI 功能未啟用</div>
          <p>請設置環境變數 GEMINI_API_KEY 後重啟服務：</p>
          <code>[Environment]::SetEnvironmentVariable("GEMINI_API_KEY", "AIzaSy...", "User")</code>
        </div>

        <div class="note note-info">
          <div class="note-title">導入模式</div>
          <div class="radio-row">
            <label><input type="radio" v-model="aiImportMode" value="library">僅導入物料庫</label>
            <label><input type="radio" v-model="aiImportMode" value="customer">導入+客戶報價單</label>
          </div>
          <div v-if="aiImportMode==='customer'" style="margin-top:10px;padding-top:10px;border-top:1px solid #C4CFE6">
            <select v-model.number="aiCustomerId" class="inp">
              <option :value="0">-- 選擇客戶 --</option>
              <option v-for="c in customerList" :key="c.id" :value="c.id">{{ c.company_name }}</option>
            </select>
          </div>
        </div>

        <div class="dropzone">
          <input type="file" ref="aiFileInput" @change="handleAiFileSelect" accept="image/*,.pdf" class="hidden">
          <div v-if="!aiFile">
            <p style="margin-bottom:10px">上傳報價單圖片或 PDF（支援手寫、掃描、多語言）</p>
            <button @click="$refs.aiFileInput.click()" :disabled="!aiStatus.available" class="btn btn-primary">選擇文件</button>
          </div>
          <div v-else>
            <p class="fname num">{{ aiFile.name }} ({{ (aiFile.size/1024).toFixed(1) }} KB)</p>
            <div class="dz-actions">
              <button @click="parseWithAI" :disabled="aiParsing || !aiStatus.available" class="btn btn-primary">
                {{ aiParsing ? 'AI 識別中...' : '開始 AI 識別' }}
              </button>
              <button @click="aiFile=null; aiProducts=[]" class="btn">重選</button>
            </div>
            <p v-if="aiParsing" class="hint">若遇伺服器繁忙，系統會自動重試，請耐心等待...</p>
          </div>
        </div>

        <div v-if="aiProducts.length > 0" class="result-wrap">
          <div class="result-bar">
            <div class="stats">
              <span>識別 {{ aiProducts.length }} 項</span>
              <span v-if="aiMeta.confidence" class="m num">置信度 {{ (aiMeta.confidence*100).toFixed(0) }}%</span>
              <span v-if="aiMeta.detected_customer" class="m">客戶：{{ aiMeta.detected_customer }}</span>
              <span v-if="aiMeta.from_cache" class="m">快取命中</span>
            </div>
            <button @click="toggleAllAIItems" class="btn btn-sm">{{ allAIConfirmed ? '取消全選' : '全選' }}</button>
          </div>
          <div class="tbl-wrap">
            <table class="tbl">
              <thead>
                <tr>
                  <th class="c" style="width:40px"><input type="checkbox" :checked="allAIConfirmed" @change="toggleAllAIItems"></th>
                  <th>型號</th><th>規格</th><th>單位</th>
                  <th class="c">數量</th><th class="r">單價</th><th>備註</th>
                </tr>
              </thead>
              <tbody>
                <tr v-for="(p, i) in aiProducts" :key="i" :class="[p.confirmed ? '' : 'row-off']">
                  <td class="c"><input type="checkbox" v-model="p.confirmed"></td>
                  <td><input v-model="p.product_code" class="cell-inp key" style="font-weight:600"></td>
                  <td><input v-model="p.product_spec" class="cell-inp"></td>
                  <td><input v-model="p.unit" class="cell-inp tc"></td>
                  <td class="c"><input type="number" v-model.number="p.quantity" class="cell-inp sm tc num"></td>
                  <td class="r"><input type="number" v-model.number="p.unit_price" class="cell-inp md tr num" style="font-weight:600"></td>
                  <td><input v-model="p.note" class="cell-inp"></td>
                </tr>
              </tbody>
            </table>
          </div>
          <div class="modal-foot">
            <button @click="closeAiModal" class="btn">取消</button>
            <button @click="submitAIImport" :disabled="aiImporting" class="btn btn-gold">
              {{ aiImporting ? '導入中...' : '確認導入' }}
            </button>
          </div>
        </div>
      </div>
    </div>
  </div>

  <!-- ============ Excel 智能導入 ============ -->
  <div v-show="showExcelModal" class="no-print modal-mask">
    <div class="modal modal-xl">
      <div class="modal-head">
        <span class="ttl">Excel 智能導入</span>
        <button @click="showExcelModal=false" class="icon-btn" title="關閉">✕</button>
      </div>
      <div class="modal-body">
        <div class="note note-info">
          <div class="radio-row">
            <label><input type="radio" v-model="excelImportMode" value="library">僅導入物料庫</label>
            <label><input type="radio" v-model="excelImportMode" value="customer">導入+客戶報價單</label>
          </div>
          <div v-if="excelImportMode==='customer'" style="margin-top:10px;padding-top:10px;border-top:1px solid #C4CFE6">
            <select v-model.number="excelCustomerId" class="inp">
              <option :value="0">-- 選擇客戶 --</option>
              <option v-for="c in customerList" :key="c.id" :value="c.id">{{ c.company_name }}</option>
            </select>
          </div>
        </div>
        <div class="dropzone">
          <input type="file" ref="excelFileInput" @change="handleExcelFileSelect" accept=".xlsx,.xls,.csv" class="hidden">
          <div v-if="!excelFile">
            <button @click="$refs.excelFileInput.click()" class="btn btn-primary">選擇 Excel 文件</button>
          </div>
          <div v-else>
            <p class="fname">{{ excelFile.name }}</p>
            <div class="dz-actions">
              <button @click="parseExcelFile" :disabled="excelParsing" class="btn btn-primary">
                {{ excelParsing ? '解析中...' : '解析文件' }}
              </button>
            </div>
          </div>
        </div>
        <div v-if="excelProducts.length > 0" class="result-wrap">
          <div class="note note-ok">識別 {{ excelProducts.length }} 項</div>
          <div class="tbl-wrap">
            <table class="tbl">
              <thead>
                <tr>
                  <th class="c" style="width:40px"><input type="checkbox" :checked="allExcelConfirmed" @change="toggleAllExcelItems"></th>
                  <th>型號</th><th>規格</th><th>單位</th>
                  <th class="c">數量</th><th class="r">單價</th><th>備註</th>
                </tr>
              </thead>
              <tbody>
                <tr v-for="(p, i) in excelProducts" :key="i" :class="[p.confirmed ? '' : 'row-off']">
                  <td class="c"><input type="checkbox" v-model="p.confirmed"></td>
                  <td><input v-model="p.product_code" class="cell-inp"></td>
                  <td><input v-model="p.product_spec" class="cell-inp"></td>
                  <td><input v-model="p.unit" class="cell-inp tc"></td>
                  <td class="c"><input type="number" v-model.number="p.quantity" class="cell-inp sm tc num"></td>
                  <td class="r"><input type="number" v-model.number="p.unit_price" class="cell-inp md tr num"></td>
                  <td><input v-model="p.note" class="cell-inp"></td>
                </tr>
              </tbody>
            </table>
          </div>
          <div class="modal-foot">
            <button @click="showExcelModal=false" class="btn">取消</button>
            <button @click="submitExcelImport" :disabled="excelImporting" class="btn btn-gold">
              {{ excelImporting ? '導入中...' : '確認導入' }}
            </button>
          </div>
        </div>
      </div>
    </div>
  </div>

  <!-- ============ CRM 客戶管理 ============ -->
  <div v-show="showCustomerModal" class="no-print modal-mask">
    <div class="modal modal-xl">
      <div class="modal-head">
        <span class="ttl">CRM 客戶管理（共 {{ customerList.length }} 筆）</span>
        <button @click="closeCustomerModal" class="icon-btn" title="關閉">✕</button>
      </div>
      <div class="modal-body">
        <div v-if="viewingCustomerQuotesFor" class="cust-orders">
          <div class="hd">
            <span>【{{ viewingCustomerQuotesFor.company_name }}】的歷史訂單</span>
            <button @click="viewingCustomerQuotesFor=null" class="link">返回 ✕</button>
          </div>
          <div class="tbl-wrap">
            <table class="tbl">
              <thead>
                <tr><th>單號</th><th>階梯</th><th class="r">含稅總額</th><th class="c">操作</th></tr>
              </thead>
              <tbody>
                <tr v-for="q in customerQuotesList" :key="q.id">
                  <td class="key num">{{ q.quote_no }}</td>
                  <td>{{ q.customer_tier }}</td>
                  <td class="r num" style="font-weight:700">{{ formatCurrency(q.final_total_amount) }} {{ q.quote_currency }}</td>
                  <td class="c">
                    <button @click="loadQuoteFromCRM(q.id)" class="btn btn-sm btn-primary">調出維護</button>
                  </td>
                </tr>
                <tr v-if="customerQuotesList.length===0"><td colspan="4" class="empty-cell">無訂單記錄</td></tr>
              </tbody>
            </table>
          </div>
        </div>
        <div v-else class="cust-form">
          <div class="hd">
            <span>{{ editingCustomer.id ? '編輯客戶' : '新增客戶' }}</span>
            <button v-if="editingCustomer.id" @click="resetCustomerForm" class="link">切換為新增</button>
          </div>
          <div class="cust-grid">
            <input v-model="editingCustomer.company_name" placeholder="* 客戶公司全稱" class="inp inp-strong">
            <input v-model="editingCustomer.tax_code" placeholder="稅號" class="inp num">
            <input v-model="editingCustomer.contact_person" placeholder="聯絡人" class="inp">
            <input v-model="editingCustomer.phone" placeholder="電話" class="inp">
            <input v-model="editingCustomer.email" placeholder="Email" class="inp">
            <select v-model="editingCustomer.default_tier" class="inp">
              <option>大型經銷商</option><option>中型經銷商</option>
              <option>中型安裝商</option><option>小型安裝商</option>
            </select>
            <input v-model="editingCustomer.address" placeholder="地址" class="inp span-2">
            <div>
              <button @click="submitSaveCustomer" class="btn btn-gold" style="width:100%">{{ editingCustomer.id ? '保存修改' : '立即新增' }}</button>
            </div>
          </div>
        </div>
        <div class="tbl-wrap">
          <table class="tbl">
            <thead>
              <tr><th>公司名稱</th><th>稅號</th><th>聯絡人</th><th>階梯</th><th class="c">訂單</th><th class="c" style="width:170px">操作</th></tr>
            </thead>
            <tbody>
              <tr v-for="c in customerList" :key="c.id">
                <td class="key">{{ c.company_name }}</td>
                <td class="dim num">{{ c.tax_code || '-' }}</td>
                <td>{{ c.contact_person }} ({{ c.phone || '-' }})</td>
                <td><span class="chip">{{ c.default_tier }}</span></td>
                <td class="c">
                  <button @click="openCustomerQuotes(c)" :class="['btn btn-sm', c.quote_count>0 ? 'btn-primary' : '']">
                    訂單 ({{ c.quote_count || 0 }})
                  </button>
                </td>
                <td class="c">
                  <div class="row-actions">
                    <button @click="selectCustomer(c); closeCustomerModal()" class="btn btn-sm btn-gold">帶入</button>
                    <button @click="editCustomer(c)" class="btn btn-sm">改</button>
                    <button @click="deleteCustomer(c.id)" class="btn btn-sm btn-danger">刪</button>
                  </div>
                </td>
              </tr>
            </tbody>
          </table>
        </div>
        <div class="modal-foot">
          <button @click="closeCustomerModal" class="btn">關閉</button>
        </div>
      </div>
    </div>
  </div>

  <!-- ============ 產品庫 ============ -->
  <div v-show="showSearchModal" class="no-print modal-mask">
    <div class="modal modal-lg">
      <div class="modal-head">
        <span class="ttl">產品庫 (共 {{ products.length }} 款)</span>
        <button @click="showSearchModal=false" class="icon-btn" title="關閉">✕</button>
      </div>
      <div class="modal-body">
        <div style="display:flex;gap:8px">
          <input v-model="modalKeyword" ref="modalInput" placeholder="輸入型號、規格、品牌..." class="inp" style="flex:1">
          <select v-model="selectedCategoryFilter" class="inp" style="width:180px">
            <option value="">全部類別</option>
            <option v-for="cat in availableCategories" :key="cat" :value="cat">{{ cat }}</option>
          </select>
        </div>
        <div class="tbl-wrap">
          <table class="tbl">
            <thead>
              <tr><th>類別</th><th>型號</th><th>規格</th><th class="r">底價</th><th class="c">操作</th></tr>
            </thead>
            <tbody>
              <tr v-for="p in searchResults" :key="p.id">
                <td class="dim">{{ p.category }}</td>
                <td class="key">{{ p.model }}</td>
                <td class="dim" style="font-size:11px">{{ p.spec }}</td>
                <td class="r num" style="font-weight:700">{{ formatCurrency(getBasePrice(p)) }}</td>
                <td class="c">
                  <button @click="addSpecificProduct(p)" class="btn btn-sm btn-primary">+ 選入</button>
                </td>
              </tr>
            </tbody>
          </table>
        </div>
        <div class="modal-foot">
          <button @click="showSearchModal=false" class="btn">關閉</button>
        </div>
      </div>
    </div>
  </div>

  <!-- ============ 歷史單據庫 ============ -->
  <div v-show="showHistoryModal" class="no-print modal-mask">
    <div class="modal modal-xl">
      <div class="modal-head">
        <span class="ttl">歷史單據庫 ({{ historyList.length }})</span>
        <button @click="showHistoryModal=false" class="icon-btn" title="關閉">✕</button>
      </div>
      <div class="modal-body">
        <div class="tbl-wrap">
          <table class="tbl">
            <thead>
              <tr><th class="c">ID</th><th>單號</th><th>客戶</th><th class="r">含稅總額</th><th class="c">狀態</th><th class="c">時間</th><th class="c">操作</th></tr>
            </thead>
            <tbody>
              <tr v-for="q in historyList" :key="q.id">
                <td class="c dim num">{{ q.id }}</td>
                <td class="key num">{{ q.quote_no }}</td>
                <td>{{ q.customer_name }}</td>
                <td class="r num" style="font-weight:700">{{ formatCurrency(q.final_total_amount) }} {{ q.quote_currency }}</td>
                <td class="c"><span class="chip chip-pos">{{ q.status }}</span></td>
                <td class="c dim num" style="font-size:11px">{{ q.created_at }}</td>
                <td class="c">
                  <button @click="loadQuoteDetail(q.id)" class="btn btn-sm btn-primary">調出維護</button>
                </td>
              </tr>
              <tr v-if="historyList.length===0"><td colspan="7" class="empty-cell">暫無歷史記錄</td></tr>
            </tbody>
          </table>
        </div>
        <div class="modal-foot">
          <button @click="showHistoryModal=false" class="btn">關閉</button>
        </div>
      </div>
    </div>
  </div>

</div>

<script>
const { createApp } = Vue;

createApp({
  data() {
    return {
      currentLang: 'vi',
      viewMode: 'customer',
      printDensity: 'auto',
      products: [],
      rates: {},
      currentQuoteId: null,
      todayDate: '',
      validUntilDate: '',
      showSearchModal: false,
      showCustomerModal: false,
      showHistoryModal: false,
      showCustomerDropdown: false,
      showAiModal: false,
      showExcelModal: false,

      aiStatus: { available: false, model: '', supported_formats: [] },
      aiFile: null, aiParsing: false, aiImporting: false,
      aiProducts: [], aiMeta: {},
      aiImportMode: 'library', aiCustomerId: 0,

      excelFile: null, excelParsing: false, excelImporting: false,
      excelProducts: [], excelMeta: {},
      excelImportMode: 'library', excelCustomerId: 0,

      viewingCustomerQuotesFor: null,
      customerQuotesList: [],
      customerList: [],
      editingCustomer: { id: null, company_name: '', tax_code: '', contact_person: '', phone: '', email: '', address: '', default_tier: '中型經銷商' },

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
        quote_no: '', layout_template_id: 1, customer_id: 0,
        customer_name: '', contact_person: '', phone_email: '',
        sales_rep: 'Thanh Bình 清平', customer_tier: '中型經銷商',
        delivery_scenario: '越南本地倉', trade_term: 'DAP',
        target_currency: 'VND', vat_rate_pct: 10.0,
        valid_days: 3, markup_pct: 0.0, lang_mode: 'vi_zh', template_style: 'standard'
      },

      termsList: [
        'Báo giá có hiệu lực trong vòng 3 ngày / 本報價有效期為 3 天。',
        'Thời gian giao hàng: 3-7 ngày / 交貨期：3-7 天內。',
        'Giá chưa bao gồm vận chuyển và lắp đặt / 單價不含運輸及安裝。',
        'Địa điểm giao hàng: Kho khách hàng chỉ định / 交貨地點：客戶指定倉庫。',
        'Thanh toán: cọc 50% khi xác nhận, 50% trước khi giao hàng / 付款：確認預付50%，發貨前付清50%。'
      ],
      tableRows: []
    };
  },
  computed: {
    t() {
      const dict = {
        vi: { workspace: 'Bàn làm việc Báo giá Nhanh', customer_view: 'Khách hàng', finance_view: 'Tài chính Nội bộ' },
        zh: { workspace: '智慧快速報價工作台', customer_view: '客戶版', finance_view: '財務版' },
        en: { workspace: 'Smart Quotation Workspace', customer_view: 'Customer', finance_view: 'Finance' }
      };
      return (k) => (dict[this.currentLang] || dict.vi)[k] || k;
    },
    allAIConfirmed() { return this.aiProducts.length > 0 && this.aiProducts.every(p => p.confirmed); },
    allExcelConfirmed() { return this.excelProducts.length > 0 && this.excelProducts.every(p => p.confirmed); },
    matchedCustomers() {
      if (!this.form.customer_name) return this.customerList.slice(0, 5);
      const kw = this.form.customer_name.trim().toLowerCase();
      return this.customerList.filter(c => c.company_name.toLowerCase().includes(kw)).slice(0, 6);
    },
    hasAnyCustomMarkup() { return this.tableRows.some(r => r.custom_markup !== null && r.custom_markup !== undefined && r.custom_markup !== ''); },
    availableCategories() { return Array.from(new Set(this.products.map(p => p.category).filter(Boolean))); },
    searchResults() {
      return this.products.filter(p => {
        const kw = this.modalKeyword.trim().toLowerCase();
        const mKw = !kw || (p.model && p.model.toLowerCase().includes(kw)) || (p.spec && p.spec.toLowerCase().includes(kw)) || (p.brand && p.brand.toLowerCase().includes(kw));
        const mCat = !this.selectedCategoryFilter || p.category === this.selectedCategoryFilter;
        return mKw && mCat;
      });
    },
    filteredProductList() {
      if (!this.quickSearchKeyword) return this.products;
      const kw = this.quickSearchKeyword.trim().toLowerCase();
      return this.products.filter(p => (p.model && p.model.toLowerCase().includes(kw)) || (p.spec && p.spec.toLowerCase().includes(kw)) || (p.brand && p.brand.toLowerCase().includes(kw)));
    },
    globalMarkupFactor() { const v = Number(this.form.markup_pct); return isNaN(v) ? 1.0 : 1.0 + v/100.0; },
    totalBeforeTax() { return this.tableRows.reduce((s, r) => s + this.calcRowSubtotal(r), 0); },
    vatAmount() { const vat = this.totalBeforeTax * (Number(this.form.vat_rate_pct || 0) / 100); return this.form.target_currency === 'VND' ? Math.round(vat) : Number(vat.toFixed(2)); },
    finalTotalAfterTax() { return this.totalBeforeTax + this.vatAmount; },
    totalBaseCost() { return this.tableRows.reduce((s, r) => s + (Number(r.quantity)||0) * (Number(r.base_price)||0), 0); },
    totalGrossProfit() { return this.totalBeforeTax - this.totalBaseCost; },
    grossMarginPercent() { return this.totalBeforeTax <= 0 ? 0 : (this.totalGrossProfit / this.totalBeforeTax) * 100; }
  },
  async mounted() {
    this.initDate();
    await Promise.all([this.loadProducts(), this.loadRates(), this.loadSettings(), this.loadCustomers(), this.loadHistory(), this.loadAIStatus()]);
    document.addEventListener('click', (e) => { if (!e.target.closest('.relative')) this.showCustomerDropdown = false; });
  },
  methods: {
    setLanguage(lang) {
      this.currentLang = lang;
      if (lang === 'en') { this.layoutConfig.title_text = 'QUOTATION 報價單'; }
      else { this.layoutConfig.title_text = 'BẢNG BÁO GIÁ 商業報價單'; }
    },
    initDate() {
      const now = new Date();
      const d = String(now.getDate()).padStart(2,'0'), m = String(now.getMonth()+1).padStart(2,'0'), y = now.getFullYear();
      this.todayDate = `${d}.${m}.${y}`;
      if (!this.form.quote_no) this.form.quote_no = `BG-${y}.${m}.${d}-001`;
      const vd = new Date(); vd.setDate(now.getDate() + 3);
      this.validUntilDate = `${String(vd.getDate()).padStart(2,'0')}.${String(vd.getMonth()+1).padStart(2,'0')}.${vd.getFullYear()}`;
    },
    async loadProducts() {
      try { const r = await fetch('/api/products'); this.products = await r.json(); if (this.products.length && !this.tableRows.length) this.addSpecificProduct(this.products[0]); } catch(e){}
    },
    async loadRates() { try { const r = await fetch('/api/rates'); this.rates = await r.json(); } catch(e){} },
    async loadSettings() {
      try {
        const r = await fetch('/api/system/global_settings'); const d = await r.json();
        if (d && d.logo_data) this.companyInfo.logo_data = d.logo_data;
      } catch(e){}
    },
    async loadCustomers() { try { const r = await fetch('/api/customers/list'); this.customerList = await r.json(); } catch(e){} },
    async loadHistory() { try { const r = await fetch('/api/quotes/list'); this.historyList = await r.json(); } catch(e){} },
    async loadAIStatus() { try { const r = await fetch('/api/ai/status'); this.aiStatus = await r.json(); } catch(e){} },
    async fetchNextQuoteNo() {
      try { const r = await fetch('/api/quotes/generate_next_no'); const d = await r.json(); if (d.next_quote_no) this.form.quote_no = d.next_quote_no; } catch(e){}
    },
    async handleLogoUpload(e) {
      const f = e.target.files[0]; if (!f) return;
      const reader = new FileReader();
      reader.onload = async (ev) => {
        this.companyInfo.logo_data = ev.target.result;
        await fetch('/api/system/save_global_logo', { method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({logo_data: ev.target.result}) });
      };
      reader.readAsDataURL(f);
    },

    openAiImportModal() { this.showAiModal = true; this.loadAIStatus(); this.loadCustomers(); },
    closeAiModal() { this.showAiModal = false; this.aiFile = null; this.aiProducts = []; this.aiMeta = {}; },
    handleAiFileSelect(e) { this.aiFile = e.target.files[0]; this.aiProducts = []; this.aiMeta = {}; },
    async parseWithAI() {
      if (!this.aiFile) return;
      this.aiParsing = true;
      try {
        const fd = new FormData(); fd.append('file', this.aiFile);
        const r = await fetch('/api/ai/parse_file', { method: 'POST', body: fd });
        const d = await r.json();
        if (r.ok && d.status === 'success') {
          this.aiProducts = (d.data.products || []).map(p => ({ ...p, confirmed: true }));
          this.aiMeta = d.data;
          if (this.aiProducts.length === 0) alert('⚠️ AI 未識別到產品，請檢查圖片清晰度');
        } else {
          alert('❌ AI 識別失敗：' + (d.detail || '未知錯誤'));
        }
      } catch(e) { alert('❌ 請求失敗：' + e); } finally { this.aiParsing = false; }
    },
    toggleAllAIItems() { const nv = !this.allAIConfirmed; this.aiProducts.forEach(p => p.confirmed = nv); },
    async submitAIImport() {
      const items = this.aiProducts.filter(p => p.confirmed);
      if (!items.length) return alert('⚠️ 請至少勾選一項');
      if (this.aiImportMode === 'customer' && !this.aiCustomerId) return alert('⚠️ 請選擇客戶');
      const cust = this.customerList.find(c => c.id === this.aiCustomerId);
      this.aiImporting = true;
      try {
        const r = await fetch('/api/import/unified', {
          method: 'POST', headers: {'Content-Type':'application/json'},
          body: JSON.stringify({
            file_name: this.aiFile.name, import_mode: this.aiImportMode, import_source: 'ai',
            ai_model: this.aiStatus.model, ai_confidence: this.aiMeta.confidence || 0,
            customer_id: this.aiCustomerId, customer_name: cust ? cust.company_name : '',
            items: items.map(p => ({ product_code: p.product_code, product_spec: p.product_spec, unit: p.unit, quantity: Number(p.quantity), unit_price: Number(p.unit_price), note: p.note, confirmed: true }))
          })
        });
        const d = await r.json();
        if (r.ok && d.status === 'success') {
          let msg = d.message;
          if (d.quotation_no) msg += `\\n\\n📄 已生成報價單：${d.quotation_no}`;
          alert('🎉 ' + msg);
          await Promise.all([this.loadProducts(), this.loadCustomers(), this.loadHistory()]);
          if (d.quotation_id) await this.loadQuoteDetail(d.quotation_id);
          this.closeAiModal();
        } else alert('❌ 導入失敗：' + (d.detail || '未知'));
      } finally { this.aiImporting = false; }
    },

    openExcelImportModal() { this.showExcelModal = true; this.loadCustomers(); },
    handleExcelFileSelect(e) { this.excelFile = e.target.files[0]; this.excelProducts = []; },
    async parseExcelFile() {
      if (!this.excelFile) return;
      this.excelParsing = true;
      try {
        const fd = new FormData(); fd.append('file', this.excelFile);
        const r = await fetch('/api/excel/parse', { method: 'POST', body: fd });
        const d = await r.json();
        if (d.status === 'success') {
          this.excelProducts = (d.products || []).map(p => ({ ...p, confirmed: true }));
          this.excelMeta = d;
        } else alert('❌ 解析失敗：' + (d.detail || '未知'));
      } finally { this.excelParsing = false; }
    },
    toggleAllExcelItems() { const nv = !this.allExcelConfirmed; this.excelProducts.forEach(p => p.confirmed = nv); },
    async submitExcelImport() {
      const items = this.excelProducts.filter(p => p.confirmed);
      if (!items.length) return alert('⚠️ 請至少勾選一項');
      if (this.excelImportMode === 'customer' && !this.excelCustomerId) return alert('⚠️ 請選擇客戶');
      const cust = this.customerList.find(c => c.id === this.excelCustomerId);
      this.excelImporting = true;
      try {
        const r = await fetch('/api/import/unified', {
          method: 'POST', headers: {'Content-Type':'application/json'},
          body: JSON.stringify({
            file_name: this.excelFile.name, import_mode: this.excelImportMode, import_source: 'excel',
            customer_id: this.excelCustomerId, customer_name: cust ? cust.company_name : '',
            items: items.map(p => ({ product_code: p.product_code, product_spec: p.product_spec, unit: p.unit, quantity: Number(p.quantity), unit_price: Number(p.unit_price), note: p.note, confirmed: true }))
          })
        });
        const d = await r.json();
        if (r.ok && d.status === 'success') {
          alert('🎉 ' + d.message);
          await Promise.all([this.loadProducts(), this.loadCustomers(), this.loadHistory()]);
          this.showExcelModal = false;
        } else alert('❌ 導入失敗：' + (d.detail || '未知'));
      } finally { this.excelImporting = false; }
    },

    openCustomerModal() { this.showCustomerModal = true; this.loadCustomers(); },
    closeCustomerModal() { this.showCustomerModal = false; this.viewingCustomerQuotesFor = null; },
    async openCustomerQuotes(c) {
      this.viewingCustomerQuotesFor = c;
      try { const r = await fetch(`/api/customers/${c.id}/quotes`); this.customerQuotesList = await r.json(); } catch(e){}
    },
    async loadQuoteFromCRM(id) { this.closeCustomerModal(); await this.loadQuoteDetail(id); },
    selectCustomer(c) {
      this.form.customer_id = c.id;
      this.form.customer_name = c.company_name;
      this.form.contact_person = c.contact_person || '';
      this.form.phone_email = c.phone || c.email || '';
      if (c.default_tier) this.form.customer_tier = c.default_tier;
      this.showCustomerDropdown = false;
    },
    editCustomer(c) { this.editingCustomer = { ...c }; },
    resetCustomerForm() { this.editingCustomer = { id: null, company_name: '', tax_code: '', contact_person: '', phone: '', email: '', address: '', default_tier: '中型經銷商' }; },
    async submitSaveCustomer() {
      if (!this.editingCustomer.company_name) return alert('請填寫公司全稱');
      const r = await fetch('/api/customers/save', { method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify(this.editingCustomer) });
      const d = await r.json();
      if (d.status === 'success') { alert('✅ 儲存成功'); this.resetCustomerForm(); this.loadCustomers(); }
    },
    async deleteCustomer(id) {
      if (!confirm('確定刪除該客戶？')) return;
      await fetch(`/api/customers/${id}`, { method: 'DELETE' });
      this.loadCustomers();
    },

    openSearchModal() { this.showSearchModal = true; this.modalKeyword = this.quickSearchKeyword; this.$nextTick(() => { if (this.$refs.modalInput) this.$refs.modalInput.focus(); }); },
    quickAddFirst() { if (this.filteredProductList.length) { this.addSpecificProduct(this.filteredProductList[0]); this.quickSearchKeyword = ''; } },
    resetCustomMarkups() { this.tableRows.forEach(r => r.custom_markup = null); },
    getBasePrice(p) {
      const rV = this.rates['USD_VND'] || 25967, rC = this.rates['USD_CNY'] || 6.70;
      const base = p.default_base_usd || 100;
      if (this.form.target_currency === 'VND') return Math.round(base * rV);
      if (this.form.target_currency === 'CNY') return Number((base * rC).toFixed(2));
      return Number(base.toFixed(2));
    },
    addSpecificProduct(p) {
      let unit = '台 / Bộ';
      if (p.category && p.category.includes('組件')) unit = '塊 / Tấm';
      else if (p.category && p.category.includes('逆變器')) unit = '套 / Bộ';
      this.tableRows.push({
        product_id: p.id, name: p.model, spec: p.spec || p.model, unit,
        quantity: 1, base_price: this.getBasePrice(p), custom_markup: null, note: ''
      });
    },
    addCustomRow() {
      this.tableRows.push({ product_id: 0, name: '自定義產品', spec: '', unit: '個 / Cái', quantity: 1, base_price: 0, custom_markup: null, note: '' });
    },
    removeRow(i) { this.tableRows.splice(i, 1); },
    getRowFactor(row) {
      if (row.custom_markup !== null && row.custom_markup !== undefined && row.custom_markup !== '') {
        const v = Number(row.custom_markup);
        return isNaN(v) ? 1.0 : 1.0 + v/100;
      }
      return this.globalMarkupFactor;
    },
    calcRowFinalUnitPrice(row) {
      const p = (Number(row.base_price) || 0) * this.getRowFactor(row);
      return this.form.target_currency === 'VND' ? Math.round(p) : Number(p.toFixed(2));
    },
    onManualUnitPriceChange(row, v) {
      const np = Number(v) || 0;
      if (row.product_id === 0) {
        const f = this.getRowFactor(row) || 1;
        row.base_price = this.form.target_currency === 'VND' ? Math.round(np/f) : Number((np/f).toFixed(2));
        return;
      }
      const bp = Number(row.base_price) || 0;
      if (bp > 0) row.custom_markup = Number((((np - bp) / bp) * 100).toFixed(2));
      else { row.base_price = np; row.custom_markup = null; }
    },
    calcRowSubtotal(row) {
      const s = (Number(row.quantity) || 0) * this.calcRowFinalUnitPrice(row);
      return this.form.target_currency === 'VND' ? Math.round(s) : Number(s.toFixed(2));
    },
    calcRowProfit(row) {
      const q = Number(row.quantity) || 0, up = this.calcRowFinalUnitPrice(row), bp = Number(row.base_price) || 0;
      const p = q * (up - bp);
      return this.form.target_currency === 'VND' ? Math.round(p) : Number(p.toFixed(2));
    },
    onCurrencyChange() {
      this.tableRows.forEach(r => {
        if (r.product_id > 0) {
          const p = this.products.find(x => x.id === r.product_id);
          if (p) r.base_price = this.getBasePrice(p);
        }
      });
    },
    addTermRow() { this.termsList.push('新增條款內容...'); },
    removeTermRow(i) { this.termsList.splice(i, 1); },
    moveTermUp(i) { if (i > 0) { [this.termsList[i-1], this.termsList[i]] = [this.termsList[i], this.termsList[i-1]]; } },
    moveTermDown(i) { if (i < this.termsList.length-1) { [this.termsList[i+1], this.termsList[i]] = [this.termsList[i], this.termsList[i+1]]; } },

    async handleNewQuote() {
      if ((this.tableRows.length || this.form.customer_name) && !confirm('⚠️ 確認要新建？當前內容將重置')) return;
      this.currentQuoteId = null;
      this.form.customer_id = 0; this.form.customer_name = '';
      this.form.contact_person = ''; this.form.phone_email = '';
      this.form.markup_pct = 0; this.form.quote_no = '';
      this.initDate(); this.tableRows = [];
      if (this.products.length) this.addSpecificProduct(this.products[0]);
      await this.fetchNextQuoteNo();
    },

    async saveToServer(forceNew = false) {
      if (!this.tableRows.length) return alert('⚠️ 請先添加明細');
      const payload = {
        quote_id: forceNew ? null : this.currentQuoteId,
        save_as_new: forceNew,
        layout_template_id: 1,
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
          product_id: r.product_id || 0, name: r.name, spec: r.spec, unit: r.unit,
          quantity: Number(r.quantity), base_price: Number(r.base_price),
          custom_markup: (r.custom_markup !== null && r.custom_markup !== undefined && r.custom_markup !== '') ? Number(r.custom_markup) : null,
          unit_price: this.calcRowFinalUnitPrice(r), subtotal: this.calcRowSubtotal(r), note: r.note || ''
        }))
      };
      try {
        const r = await fetch('/api/quote/save_or_update', { method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify(payload) });
        const d = await r.json();
        if (r.ok && d.status === 'success') {
          this.currentQuoteId = d.quote_id; this.form.quote_no = d.quote_no;
          alert('🎉 ' + d.message);
          await Promise.all([this.loadProducts(), this.loadCustomers(), this.loadHistory()]);
        } else alert('❌ ' + (d.detail || '保存失敗'));
      } catch(e) { alert('❌ 請求失敗：' + e); }
    },

    async loadQuoteDetail(id) {
      try {
        const r = await fetch(`/api/quotes/${id}`);
        const d = await r.json(); const h = d.header;
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
        if (h.terms_json) { try { this.termsList = JSON.parse(h.terms_json); } catch(e){} }
        if (h.custom_title) this.layoutConfig.title_text = h.custom_title;
        if (h.custom_greeting) this.layoutConfig.greeting_text = h.custom_greeting;
        if (h.company_logo_data) this.companyInfo.logo_data = h.company_logo_data;
        if (h.company_name_vi) this.companyInfo.name_vi = h.company_name_vi;
        if (h.company_name_cn) this.companyInfo.name_cn = h.company_name_cn;
        if (h.company_tax_code) this.companyInfo.tax_code = h.company_tax_code;
        if (h.company_contact_info) this.companyInfo.contact_info = h.company_contact_info;
        if (h.company_address) this.companyInfo.address = h.company_address;
        if (h.company_bank_info) this.companyInfo.bank_info = h.company_bank_info;
        this.tableRows = d.items.map(it => ({
          product_id: it.product_id, name: it.item_name || it.item_model,
          spec: it.item_spec || '', unit: it.unit || '台 / Bộ',
          quantity: it.quantity, base_price: Number(it.base_price) || Number(it.quote_unit_price),
          custom_markup: it.custom_markup !== null ? Number(it.custom_markup) : null,
          note: it.note || ''
        }));
        this.showHistoryModal = false;
        alert(`✅ 已調出報價單：${h.quote_no}`);
      } catch(e) { alert('❌ 加載失敗：' + e); }
    },
    async openHistoryModal() { this.showHistoryModal = true; await this.loadHistory(); },

    printOffer() { window.print(); },
    formatCurrency(v) {
      if (v === null || v === undefined || isNaN(v)) return '0';
      const n = Number(v);
      return this.form.target_currency === 'VND' ? n.toLocaleString('vi-VN', {maximumFractionDigits: 0}) : n.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2});
    }
  }
}).mount('#app');
</script>
</body>
</html>"""

# ================= 啟動 =================
if __name__ == '__main__':
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)