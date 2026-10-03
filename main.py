# -*- coding: utf-8 -*-
"""
==================================================================================================
项目名称 (Project): GALAXYTECK 智慧快速报价与客户关系管理系统 (企业标准网络版)
版本编号 (Version): v4.1.2 ASCII Environment Fix Edition
发布日期 (Date): 2026年10月
软件授权与智慧财产权声明 (Intellectual Property & Copyright):
    版权所有 (C) 2024-2026 CÔNG TY TNHH TẬP ĐOÀN CÔNG NGHỆ GALAXY VIỆT NAM
    (GALAXY VIETNAM TECHNOLOGY CORPORATION COMPANY LIMITED / 银河科技集团(越南)有限公司)
    企业代码 / Mã số doanh nghiệp: 2301296587
    地址 / Địa chỉ: Số 36 Yết Kiêu, Phường Kinh Bắc, Tỉnh Bắc Ninh, Việt Nam
==================================================================================================
"""

# ================= 修復 Windows 中文用戶名 / 非 ASCII 導致的 httpx header 編碼錯誤 =================
# 【關鍵】必須在所有其他 import 之前執行
import os as _os_env
import sys as _sys_env

def _force_ascii_env():
    """把所有可能被 SDK / httpx 讀取並寫入 HTTP header 的環境變數強制改成純 ASCII"""
    safe = {
        'USERNAME': 'galaxy',
        'USER': 'galaxy',
        'LOGNAME': 'galaxy',
        'COMPUTERNAME': 'GALAXY-PC',
        'USERDOMAIN': 'GALAXY',
        'USERDOMAIN_ROAMINGPROFILE': 'GALAXY',
        'HOMEDRIVE': 'C:',
        'HOMEPATH': r'\Users\galaxy',
        'USERPROFILE': r'C:\Users\galaxy',
        'HOME': r'C:\Users\galaxy',
        'APPDATA': r'C:\Users\galaxy\AppData\Roaming',
        'LOCALAPPDATA': r'C:\Users\galaxy\AppData\Local',
        'TEMP': r'C:\Users\galaxy\AppData\Local\Temp',
        'TMP': r'C:\Users\galaxy\AppData\Local\Temp',
    }
    for k, v in safe.items():
        if k in _os_env.environ:
            _os_env.environ[k] = v
    # 強制 UTF-8
    _os_env.environ['PYTHONIOENCODING'] = 'utf-8'
    _os_env.environ['PYTHONUTF8'] = '1'
    _os_env.environ.setdefault('LANG', 'en_US.UTF-8')
    _os_env.environ.setdefault('LC_ALL', 'en_US.UTF-8')

_force_ascii_env()

# 設定標準輸入輸出為 UTF-8（避免 print 時再出編碼問題）
try:
    import io as _io_env
    if hasattr(_sys_env.stdout, 'buffer'):
        _sys_env.stdout = _io_env.TextIOWrapper(_sys_env.stdout.buffer, encoding='utf-8', errors='replace', line_buffering=True)
    if hasattr(_sys_env.stderr, 'buffer'):
        _sys_env.stderr = _io_env.TextIOWrapper(_sys_env.stderr.buffer, encoding='utf-8', errors='replace', line_buffering=True)
except Exception:
    pass
# =================================================================================

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
import traceback

# ================= Gemini AI 配置（新版 SDK） =================
try:
    from google import genai
    from google.genai import types as genai_types
    GEMINI_AVAILABLE = True
except ImportError:
    GEMINI_AVAILABLE = False
    print("⚠️ google-genai 未安装，请执行：pip install google-genai pillow")

def _sanitize_ascii(s: str) -> str:
    """强制只保留 ASCII 字符，避免 httpx header 编码崩溃"""
    if not s:
        return ""
    return "".join(c for c in s if ord(c) < 128)

GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "") or ""
# ← 也可在此直接填 Key（请确保是纯 ASCII）
# GEMINI_API_KEY = "AIzaSy..."

# 强制清洗 API Key（防止复制粘贴时混入不可见非 ASCII 字符）
if GEMINI_API_KEY:
    cleaned = _sanitize_ascii(GEMINI_API_KEY.strip())
    if cleaned != GEMINI_API_KEY.strip():
        print("⚠️ 检测到 GEMINI_API_KEY 含有非 ASCII 字符，已自动清洗")
    GEMINI_API_KEY = cleaned

# ✅ 您的专属模型
GEMINI_MODEL = "gemini-3.8-flash"

# 新版 SDK 用 Client 实例（显式指定纯 ASCII headers，杜绝环境变量污染）
gemini_client = None
if GEMINI_AVAILABLE and GEMINI_API_KEY:
    try:
        _safe_headers = {
            "User-Agent": "GalaxyTeck-Quote/4.1.2",
            "x-goog-api-client": "google-genai-sdk",
        }
        gemini_client = genai.Client(
            api_key=GEMINI_API_KEY,
            http_options=genai_types.HttpOptions(headers=_safe_headers)
        )
        print(f"✅ Gemini AI 已启用（新版 SDK），模型：{GEMINI_MODEL}")
    except Exception as e:
        print(f"⚠️ Gemini 初始化失败：{e}")
        # 再尝试不带 http_options 的兜底
        try:
            gemini_client = genai.Client(api_key=GEMINI_API_KEY)
            print(f"✅ Gemini AI 已启用（兜底模式），模型：{GEMINI_MODEL}")
        except Exception as e2:
            print(f"⚠️ Gemini 兜底初始化也失败：{e2}")
            GEMINI_AVAILABLE = False
else:
    if GEMINI_AVAILABLE:
        print("⚠️ 未设置 GEMINI_API_KEY，AI 功能不可用（Excel 导入仍可正常使用）")


# ================= httpx header 安全补丁（最后防线） =================
# httpx 默认用 ascii 编码 header，任何非 ASCII 都会崩溃。
# 此补丁在真正写出前把非 ASCII 字符安全丢弃，避免整次请求失败。
try:
    import httpx as _httpx_patch
    _orig_norm = _httpx_patch._models._normalize_header_value
    def _safe_header_value(value, encoding=None):
        if isinstance(value, str):
            value = "".join(c for c in value if ord(c) < 128)
        return _orig_norm(value, encoding)
    _httpx_patch._models._normalize_header_value = _safe_header_value
    print("✅ httpx header 安全补丁已启用")
except Exception as _patch_err:
    print(f"⚠️ httpx 补丁未启用（可忽略）: {_patch_err}")
# =====================================================================

app = FastAPI(
    title="GalaxyTeck 智慧快速报价系统",
    version="4.1.2",
    description="企业级专业报价软件：Gemini 3.8 AI 多模态导入、三语切换、CRM 订单穿透、A4防跑版"
)

# ================= 1. 数据库连接 =================
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
                    `company_name_cn` VARCHAR(255) DEFAULT '银河科技集团（越南）有限公司',
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
                    `name_cn` VARCHAR(255) DEFAULT '银河科技集团（越南）有限公司',
                    `tax_code` VARCHAR(50) DEFAULT '2301296587',
                    `contact_info` VARCHAR(255) DEFAULT 'Điện thoại: 0325.609.218 | Email: galaxyteck6886@gmail.com',
                    `address` VARCHAR(255) DEFAULT 'Số 36 Yết Kiêu, Phường Kinh Bắc, Tỉnh Bắc Ninh, Việt Nam',
                    `bank_info` VARCHAR(255) DEFAULT 'STK (Techcombank): 314431 - Ngân hàng TMCP Kỹ thương Việt Nam',
                    `title_text` VARCHAR(255) DEFAULT 'BẢNG BÁO GIÁ 商业报价单',
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
                    `default_tier` VARCHAR(50) DEFAULT '中型经销商',
                    `payment_terms` VARCHAR(255) DEFAULT 'TT 50% 预付, 50% 发货前',
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
                ("customer_tier", "VARCHAR(50) DEFAULT '中型经销商'"),
                ("destination", "VARCHAR(100) DEFAULT '越南本地仓 (DAP)'"),
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
                ("custom_title", "VARCHAR(255) DEFAULT 'BẢNG BÁO GIÁ 商业报价单'"),
                ("custom_greeting", "TEXT DEFAULT NULL"),
                ("company_logo_data", "LONGTEXT DEFAULT NULL"),
                ("company_name_vi", "VARCHAR(255) DEFAULT 'CÔNG TY TNHH TẬP ĐOÀN CÔNG NGHỆ GALAXY VIỆT NAM'"),
                ("company_name_cn", "VARCHAR(255) DEFAULT '银河科技集团（越南）有限公司'"),
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
                ("applied_tier", "VARCHAR(50) DEFAULT '中型经销商'"),
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
        print("数据库自愈巡检完成：Gemini AI 多模态模块已 100% 准备就绪！")
    except Exception as e:
        print("数据库巡检通知:", e)

# ================= 3. 单号发号器 =================
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
        print("获取单号失败:", e)
    next_no = f"{prefix}-{str(max_seq + 1).zfill(3)}"
    return {"next_quote_no": next_no, "date_str": now.strftime('%d.%m.%Y')}

# ================= 4. 全局设置与 Logo =================
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
        print("获取全局设置失败:", e)
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

# ================= 5. 版式模板库 API =================
class LayoutTemplateModel(BaseModel):
    id: Optional[int] = None
    template_name: str
    is_default: Optional[int] = 0
    logo_data: Optional[str] = None
    name_vi: Optional[str] = "CÔNG TY TNHH TẬP ĐOÀN CÔNG NGHỆ GALAXY VIỆT NAM"
    name_cn: Optional[str] = "银河科技集团（越南）有限公司"
    tax_code: Optional[str] = "2301296587"
    contact_info: Optional[str] = "Điện thoại: 0325.609.218 | Email: galaxyteck6886@gmail.com"
    address: Optional[str] = "Số 36 Yết Kiêu, Phường Kinh Bắc, Tỉnh Bắc Ninh, Việt Nam"
    bank_info: Optional[str] = "STK (Techcombank): 314431 - Ngân hàng TMCP Kỹ thương Việt Nam"
    title_text: Optional[str] = "BẢNG BÁO GIÁ 商业报价单"
    greeting_text: Optional[str] = "Chúng tôi xin gửi đến Quý khách hàng bảng báo giá như sau / 我司现向贵司呈报以下报价表:"
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
        print("获取版式模板失败:", e)
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
            return {"status": "success", "id": tid, "message": "版式模板保存成功！"}
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
            return {"status": "success", "message": "版式模板已删除！"}
    except Exception as e:
        if conn: conn.close()
        raise HTTPException(status_code=500, detail=str(e))

# ================= 6. CRM 客户管理 API =================
class CustomerModel(BaseModel):
    id: Optional[int] = None
    company_name: str
    short_name: Optional[str] = ""
    tax_code: Optional[str] = ""
    contact_person: Optional[str] = ""
    phone: Optional[str] = ""
    email: Optional[str] = ""
    address: Optional[str] = ""
    default_tier: Optional[str] = "中型经销商"
    payment_terms: Optional[str] = "TT 50% 预付, 50% 发货前"
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
        print("获取客户列表失败:", e)
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
        print("获取客户订单失败:", e)
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
            return {"status": "success", "id": cid, "message": "客户档案保存成功！"}
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
            return {"status": "success", "message": "客户删除成功！"}
    except Exception as e:
        if conn: conn.close()
        raise HTTPException(status_code=500, detail=str(e))

# ================= 7. 产品、汇率、报价单 API =================
@app.get("/api/products")
def get_products():
    try:
        conn = get_db()
        with conn.cursor() as cursor:
            cursor.execute("""
                SELECT p.id, p.sku, p.brand, p.category, p.model, p.spec, p.power_w, 
                       p.pcs_per_pallet, p.pcs_per_container, p.warranty_years,
                       COALESCE(
                           (SELECT unit_price FROM tb_price_matrix WHERE product_id = p.id AND customer_tier = '中型经销商' LIMIT 1),
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
        print("获取产品失败:", e)
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
        print("获取汇率失败:", e)
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
        print("获取报价列表失败:", e)
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
                raise HTTPException(status_code=404, detail="报价单不存在")
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

# ================= 8. 报价单保存 =================
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
    customer_tier: str = "中型经销商"
    delivery_scenario: str = "越南本地仓"
    trade_term: str = "DAP"
    target_currency: str = "VND"
    vat_rate_pct: float = 10.0
    valid_days: int = 3
    markup_pct: float = 0.0
    terms_text: str = ""
    terms_list: Optional[List[str]] = []
    lang_mode: str = "vi_zh"
    template_style: str = "standard"
    custom_title: Optional[str] = "BẢNG BÁO GIÁ 商业报价单"
    custom_greeting: Optional[str] = "Chúng tôi xin gửi đến Quý khách hàng bảng báo giá như sau / 我司现向贵司呈报以下报价表:"
    company_logo_data: Optional[str] = None
    company_name_vi: Optional[str] = "CÔNG TY TNHH TẬP ĐOÀN CÔNG NGHỆ GALAXY VIỆT NAM"
    company_name_cn: Optional[str] = "银河科技集团（越南）有限公司"
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
        raise HTTPException(status_code=400, detail="明细列表不可为空")

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
                if (not it.product_id or it.product_id == 0) and clean_name and clean_name != '自定义产品品名 / Description':
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
                            ) VALUES (%s, '通用自定义', '通用工程辅材', %s, %s, 1, 1, 1, 1)
                        """, (new_sku, clean_name, clean_spec))
                        it.product_id = cursor.lastrowid

                        base_dec = Decimal(str(it.base_price)) if it.base_price > 0 else Decimal("0.00")
                        if req.target_currency == "VND" and rate_usd_vnd > 0:
                            base_usd = (base_dec / rate_usd_vnd).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                        elif req.target_currency == "CNY" and rate_usd_cny > 0:
                            base_usd = (base_dec / rate_usd_cny).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                        else:
                            base_usd = base_dec.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)

                        for reg in ['越南本地仓', '中国出港']:
                            for tier in ['大型经销商', '中型经销商', '中型安装商', '小型安装商']:
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
                msg = f"原报价单 [{final_quote_no}] 已成功更新保存！"
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
                msg = f"全新报价单已成功建档保存！[单号: {final_quote_no}]"

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
        print("数据库保存异常详情:", str(e))
        raise HTTPException(status_code=500, detail=f"数据库保存出错: {str(e)}")

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
    keywords = ['stt', '序号', '产品名称', '產品名稱', 'tên sản phẩm', 'product', 'model', '型号', 'đơn giá', '单价', 'unit price']
    for idx, row in enumerate(rows[:30]):
        row_text = ' '.join([str(cell).lower() for cell in row if cell is not None])
        match_count = sum(1 for kw in keywords if kw in row_text)
        if match_count >= 2:
            return idx
    return -1

def is_summary_row(row):
    row_text = ' '.join([str(cell).upper() for cell in row if cell is not None])
    summary_keywords = ['TỔNG', 'TOTAL', 'THUẾ', 'VAT', 'SUM', '税前', '含税', '增值税', 'TỔNG GIÁ', 'TỔNG TRƯỚC', 'TỔNG SAU']
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
            if 'khách hàng' in row_text.lower() or '客户名称' in row_text:
                for cell in reversed(row):
                    if cell and str(cell).strip() and 'khách' not in str(cell).lower() and '客户' not in str(cell):
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
        print("Excel 解析失败:", str(e))
        raise HTTPException(status_code=500, detail=f"Excel 解析失败: {str(e)}")

# ================= 10. Gemini AI 多模态解析（8 次重试 + 英文提示词 + 安全檔名） =================
AI_EXTRACTION_PROMPT = """You are a professional commercial quotation data extraction assistant.
Extract all material/product items from the provided file with high accuracy.

**OUTPUT REQUIREMENTS**:
1. Return PURE JSON only, no explanation text
2. JSON structure:
{
  "detected_customer": "Customer name (if any)",
  "detected_quote_no": "Quote number (if any)",
  "detected_date": "Date (if any)",
  "detected_currency": "VND/CNY/USD (inferred)",
  "confidence": 0.95,
  "products": [
    {
      "product_code": "Product model (REQUIRED)",
      "product_spec": "Specification / quality / brand",
      "unit": "Unit",
      "quantity": 1,
      "unit_price": 0.00,
      "note": "Note"
    }
  ]
}

**EXTRACTION RULES**:
- Each row in the table = one product; product_code is required, skip rows without it
- Unit price: numbers only (strip VND, commas, spaces, etc.)
- If the same model has multiple specs (e.g., ZH 3000v / 6000v), list them separately
- Currency detection: VND typically 4-7 digits, CNY typically 2-4 digits, USD typically 1-3 digits
- confidence: your overall accuracy assessment (0-1)
- If no products found, return empty array for products

**LANGUAGE SUPPORT**: Vietnamese, Simplified Chinese, Traditional Chinese, English.
IMPORTANT: Keep the ORIGINAL text for product_code, product_spec, unit, and note fields.
DO NOT translate them. Preserve Vietnamese/Chinese characters as-is.
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
        print(f"缓存查询失败：{e}")
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
        print(f"缓存保存失败：{e}")

async def gemini_extract_from_file(file_bytes: bytes, mime_type: str) -> Dict[str, Any]:
    """调用 Gemini 3.8 Flash 从文件提取物料信息（含 8 次重试 + 编码安全）"""
    if not GEMINI_AVAILABLE or not gemini_client:
        raise HTTPException(status_code=503, detail="Gemini AI 未配置，请设置 GEMINI_API_KEY 环境变量")

    # ============ MIME 標準化為純 ASCII ============
    mime_type = (mime_type or '').strip().lower()
    allowed_mimes = ['image/jpeg', 'image/png', 'image/webp', 'image/gif', 'application/pdf']
    if mime_type not in allowed_mimes:
        if mime_type.startswith('image/'):
            mime_type = 'image/jpeg'
        else:
            mime_type = 'application/pdf'

    print(f"[Gemini-DEBUG] file_bytes 長度: {len(file_bytes)}, MIME: {mime_type}")

    # 檢查 file_bytes 型別
    if not isinstance(file_bytes, (bytes, bytearray)):
        raise HTTPException(status_code=500, detail=f"file_bytes 類型錯誤：{type(file_bytes)}")

    try:
        contents = [
            AI_EXTRACTION_PROMPT,
            genai_types.Part.from_bytes(data=file_bytes, mime_type=mime_type)
        ]
        print(f"[Gemini-DEBUG] ✅ Part 建立成功")
    except Exception as part_err:
        print(f"[Gemini-DEBUG] ❌ Part 建立失敗：{type(part_err).__name__}: {part_err}")
        print(traceback.format_exc())
        raise HTTPException(status_code=500, detail=f"準備多模態資料失敗：{str(part_err)}")

    max_attempts = 8
    last_error = None

    for attempt in range(max_attempts):
        try:
            print(f"[Gemini] 第 {attempt+1}/{max_attempts} 次尝试（模型：{GEMINI_MODEL}）...")

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
            print(f"[Gemini] ✅ 成功识别 {len(cleaned_products)} 项")
            return parsed

        except json.JSONDecodeError as e:
            raise HTTPException(status_code=500, detail=f"AI 返回格式异常：{str(e)}")

        except Exception as e:
            err_str = str(e)
            last_error = err_str
            err_type = type(e).__name__

            # 完整列印錯誤堆疊
            print("=" * 70)
            print(f"[Gemini-ERROR] 類型: {err_type}")
            print(f"[Gemini-ERROR] 訊息: {err_str}")
            print(f"[Gemini-ERROR] 完整堆疊:")
            print(traceback.format_exc())
            print("=" * 70)

            # 編碼錯誤：直接報告
            if "codec can't encode" in err_str or "ascii" in err_str.lower():
                raise HTTPException(
                    status_code=500,
                    detail=f"編碼錯誤（請查看後端終端完整堆疊定位來源）：{err_str}"
                )

            if "503" in err_str or "UNAVAILABLE" in err_str or "429" in err_str or "RESOURCE_EXHAUSTED" in err_str:
                if attempt < max_attempts - 1:
                    wait = (2 ** attempt) + _random.uniform(0, 1.5)
                    print(f"[Gemini] ⏳ 服务器繁忙，等待 {wait:.1f}s 后重试...")
                    _time.sleep(wait)
                    continue
                else:
                    raise HTTPException(
                        status_code=503,
                        detail=f"Gemini 服务器连续 {max_attempts} 次繁忙，请稍后 1-2 分钟再试。（{err_str[:150]}）"
                    )

            elif "404" in err_str or "NOT_FOUND" in err_str:
                raise HTTPException(
                    status_code=500,
                    detail=f"模型 {GEMINI_MODEL} 不可用，请确认您的 API Key 支持此模型。（{err_str[:150]}）"
                )

            else:
                raise HTTPException(status_code=500, detail=f"Gemini API 调用失败：{err_str[:200]}")

    raise HTTPException(status_code=503, detail=f"Gemini 调用失败：{last_error}")

@app.post("/api/ai/parse_file")
async def ai_parse_file(file: UploadFile = File(...)):
    try:
        content = await file.read()
        original_filename = file.filename or "unknown"

        # ============ 關鍵修改：將上傳檔名轉為純 ASCII（避免 SDK HTTP header 編碼錯誤） ============
        # 保留副檔名
        ext = ''
        if '.' in original_filename:
            ext_candidate = '.' + original_filename.rsplit('.', 1)[-1].lower()
            if ext_candidate in ['.jpg', '.jpeg', '.png', '.webp', '.gif', '.pdf']:
                ext = ext_candidate

        # 生成純 ASCII 安全檔名（不給 SDK 傳中文）
        safe_filename = f"upload_{int(_time.time())}_{uuid.uuid4().hex[:6]}{ext}"
        print(f"[AI-Upload] 原始檔名: {original_filename!r} → 安全檔名: {safe_filename!r}")

        # 用原始檔名判斷 MIME
        mime_type, _ = mimetypes.guess_type(original_filename)
        if not mime_type:
            low = original_filename.lower()
            if low.endswith('.pdf'):
                mime_type = 'application/pdf'
            elif low.endswith(('.jpg', '.jpeg')):
                mime_type = 'image/jpeg'
            elif low.endswith('.png'):
                mime_type = 'image/png'
            elif low.endswith('.webp'):
                mime_type = 'image/webp'
            elif low.endswith('.gif'):
                mime_type = 'image/gif'
            else:
                raise HTTPException(status_code=400, detail=f"无法识别文件类型：{original_filename}")

        if len(content) > 20 * 1024 * 1024:
            raise HTTPException(status_code=400, detail="文件过大，请控制在 20MB 以内")

        # 使用內容 hash 查快取
        file_hash = calculate_file_hash(content)
        cached = check_ai_cache(file_hash)
        if cached:
            cached["from_cache"] = True
            cached["file_name"] = original_filename
            return {"status": "success", "data": cached}

        # 調用 Gemini
        result = await gemini_extract_from_file(content, mime_type)
        result["from_cache"] = False
        result["file_name"] = original_filename
        result["mime_type"] = mime_type
        result["file_size"] = len(content)
        result["safe_filename"] = safe_filename
        save_ai_cache(file_hash, mime_type, result)
        return {"status": "success", "data": result}

    except HTTPException:
        raise
    except Exception as e:
        print(f"AI 解析失败：{e}")
        print(traceback.format_exc())
        raise HTTPException(status_code=500, detail=f"AI 解析失败：{str(e)}")

@app.get("/api/ai/status")
def ai_status():
    return {
        "available": GEMINI_AVAILABLE and bool(GEMINI_API_KEY) and gemini_client is not None,
        "model": GEMINI_MODEL if GEMINI_AVAILABLE else None,
        "api_key_configured": bool(GEMINI_API_KEY),
        "supported_formats": ["jpg", "jpeg", "png", "webp", "gif", "pdf"] if GEMINI_AVAILABLE else []
    }

# ================= 11. 统一导入 API =================
@app.post("/api/import/unified")
def unified_import(req: UnifiedImportRequest):
    if not req.items:
        raise HTTPException(status_code=400, detail="没有可导入的产品明细")

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
                    "AI导入" if req.import_source == "ai" else "Excel导入",
                    "通用工程辅材",
                    clean_code, clean_spec
                ))
                new_pid = cursor.lastrowid
                new_count += 1

                unit_price_vnd = Decimal(str(item.unit_price)) if item.unit_price > 0 else Decimal("0.00")
                base_usd = (unit_price_vnd / rate_usd_vnd).quantize(Decimal('0.01'), ROUND_HALF_UP) if rate_usd_vnd > 0 else Decimal("0.00")

                for reg in ['越南本地仓', '中国出港']:
                    for tier in ['大型经销商', '中型经销商', '中型安装商', '小型安装商']:
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
                    ) VALUES (%s, %s, %s, '中型经销商', '越南本地仓 (DAP)', 'VND', 0.0, 10.0, %s, 3, '已生效')
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
                        ) VALUES (%s, 0, %s, %s, %s, %s, %s, '中型经销商', %s, %s, %s, %s)
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
                "message": f"导入成功！新增 {new_count} 项，已存在 {existing_count} 项。",
                "new_count": new_count,
                "existing_count": existing_count,
                "quotation_id": quotation_id,
                "quotation_no": quotation_no,
                "imported_products": imported_products
            }
    except Exception as e:
        if conn: conn.close()
        print(f"导入失败：{e}")
        raise HTTPException(status_code=500, detail=f"导入失败：{str(e)}")

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
        print("获取导入历史失败:", e)
        return []

# ================= 12. 前端主页面（v4.1.0 UI，未改动） =================
@app.get("/", response_class=HTMLResponse)
def index():
    return """<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>GalaxyTeck 智慧快速报价系统 v4.1.2</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <script src="https://unpkg.com/vue@3/dist/vue.global.js"></script>
  <style>
    [v-cloak] { display: none !important; }

    :root {
      --fin-navy: #1E3A8A;
      --fin-navy-dark: #172554;
      --fin-navy-light: #2563EB;
      --fin-gold: #D97706;
      --fin-gold-light: #F59E0B;
      --fin-red: #DC2626;
      --fin-green: #059669;
      --fin-gray-50: #F8FAFC;
      --fin-gray-100: #F1F5F9;
      --fin-gray-200: #E2E8F0;
      --fin-gray-800: #1E293B;
    }

    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Microsoft YaHei", Roboto, Arial, sans-serif;
      background: linear-gradient(135deg, #F1F5F9 0%, #E2E8F0 100%);
      color: #1E293B;
      min-height: 100vh;
    }

    .num-font { font-family: 'Consolas', 'Monaco', 'Courier New', monospace; font-variant-numeric: tabular-nums; }

    .top-logo-img { max-height: 36px; width: auto; object-fit: contain; }
    .sheet-logo-img { max-height: 48px; width: auto; object-fit: contain; }

    .fin-card {
      background: white;
      border: 1px solid var(--fin-gray-200);
      border-radius: 12px;
      box-shadow: 0 1px 3px rgba(30, 58, 138, 0.08), 0 4px 12px -2px rgba(30, 58, 138, 0.06);
      transition: all 0.2s ease;
    }
    .fin-card:hover {
      box-shadow: 0 4px 8px rgba(30, 58, 138, 0.1), 0 8px 20px -4px rgba(30, 58, 138, 0.08);
    }

    .btn-primary {
      background: linear-gradient(135deg, #1E3A8A 0%, #2563EB 100%);
      color: white;
      font-weight: 600;
      border-radius: 10px;
      box-shadow: 0 2px 4px rgba(30, 58, 138, 0.2), 0 1px 2px rgba(30, 58, 138, 0.15);
      transition: all 0.15s ease;
    }
    .btn-primary:hover {
      background: linear-gradient(135deg, #172554 0%, #1E3A8A 100%);
      box-shadow: 0 4px 8px rgba(30, 58, 138, 0.3);
      transform: translateY(-1px);
    }
    .btn-success {
      background: linear-gradient(135deg, #059669 0%, #10B981 100%);
      color: white;
      font-weight: 600;
      box-shadow: 0 2px 4px rgba(5, 150, 105, 0.25);
    }
    .btn-success:hover {
      background: linear-gradient(135deg, #047857 0%, #059669 100%);
      transform: translateY(-1px);
    }
    .btn-ai {
      background: linear-gradient(135deg, #7C3AED 0%, #A855F7 50%, #EC4899 100%);
      color: white;
      font-weight: 700;
      box-shadow: 0 2px 6px rgba(168, 85, 247, 0.35);
      position: relative;
      overflow: hidden;
    }
    .btn-ai:hover {
      box-shadow: 0 4px 12px rgba(168, 85, 247, 0.5);
      transform: translateY(-1px);
    }

    .fin-input {
      border: 1.5px solid var(--fin-gray-200);
      border-radius: 8px;
      padding: 8px 12px;
      font-size: 13px;
      background: white;
      transition: all 0.15s ease;
      width: 100%;
    }
    .fin-input:focus {
      outline: none;
      border-color: var(--fin-navy-light);
      box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.1);
      background: #FEFEFF;
    }
    .fin-input.num-font { text-align: right; }

    .fin-thead {
      background: linear-gradient(180deg, #1E3A8A 0%, #1E40AF 100%);
      color: white;
      font-weight: 600;
      font-size: 12px;
      letter-spacing: 0.02em;
      text-shadow: 0 1px 2px rgba(0,0,0,0.15);
    }
    .fin-thead th {
      padding: 10px 8px;
      border-right: 1px solid rgba(255,255,255,0.15);
    }

    .fin-amount-gold {
      font-size: 22px;
      font-weight: 900;
      background: linear-gradient(135deg, #D97706 0%, #F59E0B 100%);
      -webkit-background-clip: text;
      -webkit-text-fill-color: transparent;
      background-clip: text;
      font-family: 'Consolas', monospace;
    }
    .fin-label {
      font-size: 11px;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.05em;
      color: #64748B;
    }

    .finance-mode-alert {
      background: linear-gradient(90deg, #FEF3C7 0%, #FED7AA 100%);
      border-left: 4px solid #DC2626;
      box-shadow: inset 0 1px 3px rgba(220, 38, 38, 0.1);
    }
    .finance-badge {
      background: linear-gradient(135deg, #DC2626 0%, #B91C1C 100%);
      color: white;
      padding: 3px 10px;
      border-radius: 6px;
      font-size: 10px;
      font-weight: 800;
      letter-spacing: 0.08em;
      box-shadow: 0 1px 3px rgba(220, 38, 38, 0.4);
    }

    @page { size: A4 portrait; margin: 12mm 14mm; }
    @media print {
      .no-print, header, nav { display: none !important; }
      body { background: white !important; font-size: 10pt !important; margin: 0 !important; padding: 0 !important; }
      .print-sheet { width: 100% !important; max-width: 100% !important; margin: 0 !important; padding: 0 !important; border: none !important; box-shadow: none !important; background: white !important; }
      input, textarea, select { border: none !important; background: transparent !important; padding: 0 !important; }
      table { page-break-inside: auto !important; width: 100% !important; border-collapse: collapse !important; }
      thead { display: table-header-group !important; }
      tr { page-break-inside: avoid !important; }
      .keep-together { page-break-inside: avoid !important; }
      .compact-mode .header-box { margin-bottom: 2mm !important; }
      .compact-mode .item-table th, .compact-mode .item-table td { padding-top: 2px !important; padding-bottom: 2px !important; }
    }

    .a4-preview-card {
      width: 210mm; min-height: 297mm; margin: 0 auto; background: white;
      box-shadow: 0 4px 6px -1px rgba(30, 58, 138, 0.08), 0 20px 40px -10px rgba(30, 58, 138, 0.12);
      border: 1px solid #CBD5E1; border-radius: 4px; padding: 12mm 14mm;
      box-sizing: border-box; position: relative;
    }

    .auto-expand-text { field-sizing: content; resize: none; }

    .fin-table-row:hover { background: linear-gradient(90deg, #F8FAFC 0%, #EFF6FF 100%); }

    ::-webkit-scrollbar { width: 8px; height: 8px; }
    ::-webkit-scrollbar-track { background: #F1F5F9; }
    ::-webkit-scrollbar-thumb { background: #CBD5E1; border-radius: 4px; }
    ::-webkit-scrollbar-thumb:hover { background: #94A3B8; }

    .badge-navy {
      background: linear-gradient(135deg, #DBEAFE 0%, #BFDBFE 100%);
      color: #1E40AF;
      border: 1px solid #93C5FD;
      padding: 3px 10px;
      border-radius: 20px;
      font-size: 11px;
      font-weight: 700;
    }
    .badge-gold {
      background: linear-gradient(135deg, #FEF3C7 0%, #FDE68A 100%);
      color: #92400E;
      border: 1px solid #FCD34D;
      padding: 3px 10px;
      border-radius: 20px;
      font-size: 11px;
      font-weight: 700;
    }
  </style>
</head>
<body>
<div id="app" v-cloak class="py-4 px-2 md:px-6 max-w-[1360px] mx-auto">

  <header class="no-print fin-card p-4 mb-5 flex flex-col xl:flex-row justify-between items-center gap-3">
    <div class="flex items-center gap-3.5">
      <label class="cursor-pointer block">
        <div v-if="companyInfo.logo_data" class="h-11 px-3 py-1 border-2 border-slate-200 rounded-xl flex items-center bg-white hover:border-blue-400 transition">
          <img :src="companyInfo.logo_data" class="top-logo-img">
        </div>
        <div v-else class="h-11 px-4 text-white rounded-xl font-bold text-sm flex items-center gap-2" style="background: linear-gradient(135deg, #1E3A8A 0%, #2563EB 100%); box-shadow: 0 2px 6px rgba(30,58,138,0.3);">
          <span>🏢 GalaxyTECK</span>
          <span class="text-xs bg-white/20 px-1.5 py-0.5 rounded">换Logo</span>
        </div>
        <input type="file" @change="handleLogoUpload" accept="image/*" class="hidden">
      </label>
      <div>
        <div class="flex items-center gap-2">
          <span class="text-base font-bold" style="color: #1E3A8A;">{{ t('workspace') }}</span>
          <span class="px-2 py-0.5 badge-navy font-mono">v4.1.2</span>
        </div>
        <div class="text-[11px] text-slate-500 mt-0.5">GALAXY VIỆT NAM · MST: 2301296587</div>
      </div>
    </div>

    <div class="flex flex-wrap items-center gap-2.5 bg-slate-100 p-1.5 rounded-xl border border-slate-200">
      <div class="flex items-center bg-white rounded-lg p-0.5 shadow-sm border text-xs">
        <button @click="setLanguage('vi')" :class="['px-2.5 py-1.5 rounded font-bold text-[11px] transition', currentLang==='vi' ? 'bg-red-600 text-white' : 'text-slate-600']">🇻🇳 Vi</button>
        <button @click="setLanguage('zh')" :class="['px-2.5 py-1.5 rounded font-bold text-[11px] transition', currentLang==='zh' ? 'text-white' : 'text-slate-600']" :style="currentLang==='zh' ? 'background: linear-gradient(135deg, #1E3A8A, #2563EB)' : ''">🇨🇳 中</button>
        <button @click="setLanguage('en')" :class="['px-2.5 py-1.5 rounded font-bold text-[11px] transition', currentLang==='en' ? 'bg-emerald-600 text-white' : 'text-slate-600']">🇬🇧 En</button>
      </div>
      <div class="flex items-center bg-white rounded-lg p-0.5 shadow-sm border text-xs">
        <button @click="viewMode='customer'" :class="['px-3 py-1.5 rounded-md font-bold transition', viewMode==='customer' ? 'text-white shadow' : 'text-slate-600']" :style="viewMode==='customer' ? 'background: linear-gradient(135deg, #1E3A8A, #2563EB)' : ''">{{ t('customer_view') }}</button>
        <button @click="viewMode='finance'" :class="['px-3 py-1.5 rounded-md font-bold transition', viewMode==='finance' ? 'text-white shadow' : 'text-slate-600']" :style="viewMode==='finance' ? 'background: linear-gradient(135deg, #D97706, #F59E0B)' : ''">{{ t('finance_view') }}</button>
      </div>
    </div>

    <div class="flex flex-wrap items-center gap-2">
      <button @click="openAiImportModal" class="btn-ai px-3.5 py-2 text-xs flex items-center gap-1.5">
        <span>🤖</span><span>AI 智能导入</span>
      </button>
      <button @click="openExcelImportModal" class="btn-primary px-3.5 py-2 text-xs flex items-center gap-1.5">
        <span>📥</span><span>Excel 导入</span>
      </button>
      <button @click="openCustomerModal" class="btn-success px-3.5 py-2 text-xs flex items-center gap-1.5">
        <span>👥</span><span>CRM</span>
      </button>
      <button @click="openHistoryModal" class="px-3.5 py-2 bg-white hover:bg-slate-50 text-slate-700 rounded-xl text-xs font-semibold border-2 border-slate-200 hover:border-blue-400 transition flex items-center gap-1.5">
        <span>📂</span><span>历史 ({{ historyList.length }})</span>
      </button>
    </div>
  </header>

  <div v-if="viewMode==='finance'" class="no-print mb-4 p-3.5 rounded-xl finance-mode-alert flex justify-between items-center text-xs">
    <div class="flex items-center gap-2.5 font-bold text-amber-900">
      <span class="finance-badge">🔒 内部机密</span>
      <span>当前为【内部财务审核版】：显示成本、调价率、利润。严禁发送给客户！</span>
    </div>
    <button @click="viewMode='customer'" class="text-blue-700 hover:underline font-bold">切回客户版 →</button>
  </div>

  <div class="no-print fin-card p-5 mb-5 space-y-4">
    <div class="text-xs font-bold border-b-2 pb-3 flex justify-between items-center" style="color: #1E3A8A; border-color: #E2E8F0;">
      <span class="flex items-center gap-2">
        <span class="w-1.5 h-4 rounded" style="background: linear-gradient(180deg, #1E3A8A, #2563EB);"></span>
        商业计价与交付参数
      </span>
      <span class="num-font text-blue-700 font-bold">💱 1 USD = {{ rates['USD_VND'] || 25967 }} VND · 1 USD = {{ rates['USD_CNY'] || 6.70 }} CNY</span>
    </div>

    <div class="grid grid-cols-1 md:grid-cols-4 gap-3 text-xs">
      <div>
        <div class="flex justify-between items-center mb-1.5">
          <label class="fin-label">报价单号</label>
          <button @click="fetchNextQuoteNo" class="text-[10px] text-blue-600 hover:underline font-bold">换新单号</button>
        </div>
        <input v-model="form.quote_no" class="fin-input num-font font-bold text-blue-700" style="background: #EFF6FF;">
      </div>
      <div class="relative">
        <div class="flex justify-between items-center mb-1.5">
          <label class="fin-label">客户名称</label>
          <button @click="openCustomerModal" class="text-[10px] text-blue-600 hover:underline font-bold">CRM选客户</button>
        </div>
        <input v-model="form.customer_name" @focus="showCustomerDropdown=true" placeholder="输入或选择客户..." class="fin-input font-semibold">
        <div v-if="showCustomerDropdown && matchedCustomers.length>0" class="absolute left-0 right-0 top-full mt-1.5 bg-white border-2 border-blue-200 rounded-xl shadow-2xl z-40 max-h-52 overflow-y-auto p-1.5">
          <div v-for="c in matchedCustomers" :key="c.id" @click="selectCustomer(c)" class="p-2.5 hover:bg-blue-50 rounded-lg cursor-pointer transition">
            <div class="font-bold text-slate-800 text-xs">{{ c.company_name }}</div>
            <div class="text-[11px] text-slate-500 mt-0.5">{{ c.contact_person }} · {{ c.phone }}</div>
          </div>
        </div>
      </div>
      <div>
        <label class="block fin-label mb-1.5">联系人</label>
        <input v-model="form.contact_person" class="fin-input">
      </div>
      <div>
        <label class="block fin-label mb-1.5">电话 / 邮箱</label>
        <input v-model="form.phone_email" class="fin-input">
      </div>
    </div>

    <div class="grid grid-cols-2 md:grid-cols-5 gap-3 text-xs">
      <div>
        <label class="block fin-label mb-1.5">交付场景</label>
        <select v-model="form.delivery_scenario" class="fin-input">
          <option value="越南本地仓">越南本地仓 (DAP)</option>
          <option value="中国出港">中国出港 (FOB)</option>
        </select>
      </div>
      <div>
        <label class="block fin-label mb-1.5">价格阶梯</label>
        <select v-model="form.customer_tier" class="fin-input font-semibold">
          <option>大型经销商</option><option>中型经销商</option>
          <option>中型安装商</option><option>小型安装商</option>
        </select>
      </div>
      <div>
        <label class="block fin-label mb-1.5">结算币种</label>
        <select v-model="form.target_currency" @change="onCurrencyChange" class="fin-input font-bold text-blue-700">
          <option value="VND">VND 越南盾 ₫</option>
          <option value="CNY">CNY 人民币 ¥</option>
          <option value="USD">USD 美元 $</option>
        </select>
      </div>
      <div>
        <label class="block fin-label mb-1.5">增值税 VAT (%)</label>
        <input type="number" v-model.number="form.vat_rate_pct" class="fin-input num-font text-right">
      </div>
      <div class="p-2 rounded-xl border-2" style="background: linear-gradient(135deg, #FEF3C7 0%, #FDE68A 100%); border-color: #F59E0B;">
        <label class="block text-[11px] font-bold mb-0.5" style="color: #92400E;">⚡ 全局调价 (%)</label>
        <input type="number" step="0.5" v-model.number="form.markup_pct" class="fin-input num-font text-right font-bold" style="color: #92400E; border-color: #F59E0B;">
      </div>
    </div>

    <div class="flex flex-wrap gap-2.5 pt-3 border-t border-slate-100 items-center">
      <button @click="openSearchModal" class="btn-primary px-4 py-2 text-xs">🔍 产品库</button>
      <input v-model="quickSearchKeyword" @keyup.enter="quickAddFirst" placeholder="输入型号搜索 (Enter 快速选入)..." class="flex-1 min-w-[280px] border-2 border-slate-200 rounded-full px-4 py-2 text-xs focus:border-blue-500 outline-none transition">
      <button @click="resetCustomMarkups" v-if="hasAnyCustomMarkup" class="px-3 py-2 rounded-xl text-xs font-bold border-2" style="background: #FEF3C7; color: #92400E; border-color: #F59E0B;">↺ 恢复全局调价</button>
      <button @click="addCustomRow" class="px-4 py-2 rounded-xl text-xs font-bold text-white" style="background: linear-gradient(135deg, #475569 0%, #1E293B 100%);">+ 自定义物料</button>
    </div>
  </div>

  <div class="no-print max-w-[210mm] mx-auto mb-3 flex justify-between items-center gap-2 px-1">
    <span class="px-3 py-1.5 bg-white border-2 border-slate-200 rounded-lg num-font font-bold text-xs" style="color: #1E3A8A;">
      单号：{{ form.quote_no || '未生成' }}
    </span>
    <div class="flex gap-2">
      <button @click="handleNewQuote" class="px-3.5 py-2 rounded-xl text-xs font-bold text-white" style="background: linear-gradient(135deg, #0D9488 0%, #14B8A6 100%); box-shadow: 0 2px 4px rgba(13,148,136,0.25);">📄✨ 新建空白单</button>
      <button v-if="currentQuoteId" @click="saveToServer(true)" class="px-3.5 py-2 rounded-xl text-xs font-bold text-white" style="background: linear-gradient(135deg, #0284C7 0%, #0EA5E9 100%); box-shadow: 0 2px 4px rgba(2,132,199,0.25);">📋 另存为新单</button>
      <button @click="saveToServer(false)" class="btn-success px-3.5 py-2 text-xs">💾 {{ currentQuoteId ? '覆盖更新' : '保存入库' }}</button>
      <button @click="printOffer" class="btn-primary px-3.5 py-2 text-xs">🖨️ 打印</button>
    </div>
  </div>

  <div :class="['a4-preview-card print-sheet', printDensity === 'ultra' || tableRows.length <= 6 ? 'compact-mode' : '']">
    <div class="header-box border-b-2 pb-3 mb-3" style="border-color: #1E3A8A;">
      <div class="flex justify-between items-start gap-4">
        <div class="max-w-[240px]">
          <div v-if="companyInfo.logo_data"><img :src="companyInfo.logo_data" class="sheet-logo-img"></div>
          <div v-else>
            <span class="text-2xl font-black" style="color: #1E3A8A;">G</span><span class="text-xl font-black">GalaxyTECK</span>
          </div>
        </div>
        <div class="flex-1 text-right">
          <textarea v-model="companyInfo.name_vi" rows="1" class="auto-expand-text text-sm font-bold text-right w-full bg-transparent border-b border-transparent"></textarea>
          <textarea v-model="companyInfo.name_cn" rows="1" class="auto-expand-text text-[11px] text-right w-full bg-transparent border-b border-transparent mt-0.5"></textarea>
        </div>
      </div>
      <div class="mt-2 pt-2 border-t grid grid-cols-2 text-[10px] gap-x-4">
        <div>MST: <input v-model="companyInfo.tax_code" class="bg-transparent num-font"></div>
        <div class="text-right"><input v-model="companyInfo.contact_info" class="bg-transparent text-right w-full"></div>
        <div class="col-span-2">地址: <input v-model="companyInfo.address" class="bg-transparent w-[85%]"></div>
        <div class="col-span-2"><input v-model="companyInfo.bank_info" class="bg-transparent w-full num-font text-[9.5px]"></div>
      </div>
    </div>

    <div class="text-center my-3">
      <div v-if="viewMode==='finance'" class="inline-block finance-badge mb-1">内部审批专用 · 严禁外发</div>
      <input v-model="layoutConfig.title_text" class="text-lg font-black text-center w-full bg-transparent" style="color: #1E3A8A;">
    </div>

    <div class="grid grid-cols-2 text-[11px] mb-3 divide-x-2 border-2" style="border-color: #CBD5E1;">
      <div class="p-2.5 space-y-1">
        <div class="font-bold border-b pb-1 text-[10px] uppercase" style="color: #1E3A8A;">客户信息</div>
        <div class="flex"><span class="w-24 text-slate-500">客户：</span><input v-model="form.customer_name" class="flex-1 font-semibold bg-transparent"></div>
        <div class="flex"><span class="w-24 text-slate-500">联系人：</span><input v-model="form.contact_person" class="flex-1 bg-transparent"></div>
        <div class="flex"><span class="w-24 text-slate-500">电话：</span><input v-model="form.phone_email" class="flex-1 bg-transparent"></div>
      </div>
      <div class="p-2.5 space-y-1">
        <div class="font-bold border-b pb-1 text-[10px] uppercase" style="color: #1E3A8A;">报价信息</div>
        <div class="flex justify-between"><span class="text-slate-500">单号：</span><input v-model="form.quote_no" class="text-right font-bold num-font bg-transparent"></div>
        <div class="flex justify-between"><span class="text-slate-500">日期：</span><span class="num-font">{{ todayDate }}</span></div>
        <div class="flex justify-between"><span class="text-slate-500">报价人：</span><input v-model="form.sales_rep" class="text-right bg-transparent"></div>
        <div class="flex justify-between"><span class="text-slate-500">有效期：</span><span class="num-font font-semibold" style="color: #DC2626;">{{ validUntilDate }} ({{ form.valid_days }} 天)</span></div>
      </div>
    </div>

    <div v-if="viewMode==='finance'" class="mb-3 p-3 rounded-lg keep-together border-2" style="background: linear-gradient(135deg, #FEF3C7 0%, #FFF7ED 100%); border-color: #F59E0B;">
      <div class="font-bold text-[11px] mb-2 border-b-2 pb-1.5 flex justify-between items-center" style="color: #92400E; border-color: #FCD34D;">
        <span>📊 内部利润与成本分析</span>
        <span class="num-font text-[10px] bg-white px-2 py-0.5 rounded-full border" style="border-color: #FCD34D;">全局调价: {{ form.markup_pct }}%</span>
      </div>
      <div class="grid grid-cols-4 gap-2 text-center">
        <div class="p-2 bg-white rounded-lg border" style="border-color: #FCD34D;">
          <div class="fin-label mb-0.5">总采购成本</div>
          <div class="num-font text-xs font-bold text-slate-700">{{ formatCurrency(totalBaseCost) }}</div>
        </div>
        <div class="p-2 bg-white rounded-lg border" style="border-color: #FCD34D;">
          <div class="fin-label mb-0.5">税前营收</div>
          <div class="num-font text-xs font-bold" style="color: #1E3A8A;">{{ formatCurrency(totalBeforeTax) }}</div>
        </div>
        <div class="p-2 bg-white rounded-lg border" style="border-color: #FCD34D;">
          <div class="fin-label mb-0.5">毛利额</div>
          <div class="num-font text-sm font-black" :style="totalGrossProfit>=0 ? 'color: #059669;' : 'color: #DC2626;'">{{ formatCurrency(totalGrossProfit) }}</div>
        </div>
        <div class="p-2 bg-white rounded-lg border" style="border-color: #FCD34D;">
          <div class="fin-label mb-0.5">毛利率</div>
          <div class="num-font text-sm font-black" :style="grossMarginPercent>=0 ? 'color: #059669;' : 'color: #DC2626;'">{{ grossMarginPercent.toFixed(2) }}%</div>
        </div>
      </div>
    </div>

    <div class="text-[10px] italic mb-1.5">
      <input v-model="layoutConfig.greeting_text" class="w-full bg-transparent italic text-slate-600">
    </div>

    <table class="item-table w-full border-collapse text-[11px] mb-3 border-2" style="border-color: #94A3B8;">
      <thead class="fin-thead">
        <tr>
          <th class="w-8">STT<br>序号</th>
          <th class="text-left px-2">产品名称</th>
          <th class="text-left px-2">规格型号</th>
          <th class="w-12">单位</th>
          <th class="w-12">数量</th>
          <th v-if="viewMode==='finance'" class="text-right w-20" style="background: #172554;">底价成本</th>
          <th class="no-print text-center w-16" style="background: #92400E;">独立调价</th>
          <th class="text-right w-24" style="background: #B45309;">对外单价</th>
          <th class="text-right w-28">总价</th>
          <th v-if="viewMode==='finance'" class="text-right w-20" style="background: #047857;">毛利</th>
          <th class="no-print w-8">操作</th>
        </tr>
      </thead>
      <tbody>
        <tr v-for="(row, idx) in tableRows" :key="idx" class="fin-table-row border-b border-slate-200">
          <td class="border border-slate-300 py-1.5 px-1 text-center num-font text-slate-500">{{ idx + 1 }}</td>
          <td class="border border-slate-300 py-1 px-2"><input v-model="row.name" class="w-full bg-transparent font-semibold"></td>
          <td class="border border-slate-300 py-1 px-2"><input v-model="row.spec" class="w-full bg-transparent num-font text-[10px] text-slate-600"></td>
          <td class="border border-slate-300 py-1 px-1"><input v-model="row.unit" class="w-full text-center bg-transparent text-slate-600"></td>
          <td class="border border-slate-300 py-1 px-1"><input type="number" v-model.number="row.quantity" class="w-10 text-center num-font font-semibold bg-transparent"></td>
          <td v-if="viewMode==='finance'" class="border border-slate-300 py-1 px-1" style="background: #F8FAFC;"><input type="number" v-model.number="row.base_price" class="w-16 text-right num-font bg-transparent"></td>
          <td class="no-print border border-slate-300 py-1.5 px-1 text-center" style="background: #FFFBEB;">
            <input type="number" step="0.5" v-model.number="row.custom_markup" :placeholder="form.markup_pct" class="w-12 border-2 rounded px-0.5 text-center num-font text-[11px] font-bold bg-white" style="border-color: #FCD34D; color: #92400E;">
          </td>
          <td class="border border-slate-300 py-1 px-2" style="background: #FFFBEB;">
            <input type="number" :value="calcRowFinalUnitPrice(row)" @change="onManualUnitPriceChange(row, $event.target.value)" class="w-20 text-right num-font font-bold bg-transparent" style="color: #92400E;">
          </td>
          <td class="border border-slate-300 py-1 px-2 text-right num-font font-bold" style="color: #1E3A8A;">{{ formatCurrency(calcRowSubtotal(row)) }}</td>
          <td v-if="viewMode==='finance'" class="border border-slate-300 py-1 px-1 text-right num-font font-bold" style="color: #047857; background: #ECFDF5;">{{ formatCurrency(calcRowProfit(row)) }}</td>
          <td class="no-print border border-slate-300 py-1.5 px-1 text-center"><button @click="removeRow(idx)" class="text-red-500 hover:text-red-700 font-bold text-sm">✕</button></td>
        </tr>
        <tr v-if="tableRows.length===0">
          <td :colspan="viewMode==='finance' ? 11 : 9" class="text-center py-8 text-slate-400 text-xs">清单为空，请新增产品</td>
        </tr>
      </tbody>
    </table>

    <div class="keep-together flex justify-end mb-3">
      <div class="w-80 rounded-lg overflow-hidden border-2" style="border-color: #1E3A8A;">
        <div class="flex justify-between py-2 px-3 bg-slate-50 border-b">
          <span class="font-bold text-xs text-slate-700">税前总额：</span>
          <span class="num-font font-bold text-sm" style="color: #1E3A8A;">{{ formatCurrency(totalBeforeTax) }} {{ form.target_currency }}</span>
        </div>
        <div class="flex justify-between py-2 px-3 bg-slate-50 border-b">
          <span class="text-xs text-slate-600">增值税 VAT ({{ form.vat_rate_pct }}%)：</span>
          <span class="num-font font-bold text-sm text-slate-700">{{ formatCurrency(vatAmount) }} {{ form.target_currency }}</span>
        </div>
        <div class="flex justify-between py-3 px-3" style="background: linear-gradient(135deg, #1E3A8A 0%, #2563EB 100%);">
          <span class="font-black text-white text-sm">含税总额：</span>
          <span class="fin-amount-gold" style="color: white; -webkit-text-fill-color: white;">{{ formatCurrency(finalTotalAfterTax) }} {{ form.target_currency }}</span>
        </div>
      </div>
    </div>

    <div class="terms-box keep-together mb-4 text-[10px] p-3 rounded-lg border-2" style="background: #F8FAFC; border-color: #CBD5E1;">
      <div class="flex justify-between items-center mb-1.5 border-b pb-1.5" style="border-color: #E2E8F0;">
        <div class="font-bold text-[11px]" style="color: #1E3A8A;">商务条款及备注：</div>
        <button @click="addTermRow" class="no-print text-[10px] px-2.5 py-1 rounded font-bold text-white" style="background: #1E3A8A;">+ 新增</button>
      </div>
      <div class="space-y-1">
        <div v-for="(term, tIdx) in termsList" :key="tIdx" class="flex items-start gap-1 group">
          <span class="font-bold num-font text-[10px] pt-0.5" style="color: #1E3A8A;">{{ tIdx + 1 }})</span>
          <input v-model="termsList[tIdx]" class="flex-1 bg-transparent border-b border-transparent hover:border-slate-300 text-[10px] text-slate-700">
          <div class="no-print opacity-0 group-hover:opacity-100 flex gap-1 transition">
            <button @click="moveTermUp(tIdx)" :disabled="tIdx===0" class="text-slate-400 disabled:opacity-30">▲</button>
            <button @click="moveTermDown(tIdx)" :disabled="tIdx===termsList.length-1" class="text-slate-400 disabled:opacity-30">▼</button>
            <button @click="removeTermRow(tIdx)" class="text-rose-500 font-bold px-1">✕</button>
          </div>
        </div>
      </div>
    </div>

    <div class="signature-box keep-together grid grid-cols-2 text-center text-[11px] pt-2 border-t-2" style="border-color: #1E3A8A;">
      <div>
        <div class="font-bold mb-1" style="color: #1E3A8A;">{{ viewMode==='customer' ? '客户确认签字盖章' : '制单人' }}</div>
        <div class="h-14 flex items-end justify-center text-slate-400 italic text-[9px]">(签字盖章)</div>
      </div>
      <div>
        <div class="font-bold mb-1" style="color: #1E3A8A;">{{ viewMode==='customer' ? '卖方确认盖章' : '总经理审批' }}</div>
        <div class="h-14 flex flex-col justify-end items-center">
          <div class="font-bold text-[10px]">{{ companyInfo.name_vi }}</div>
          <div class="font-semibold text-[10px]">GIÁM ĐỐC: WANG XIAN LI</div>
        </div>
      </div>
    </div>

    <div class="keep-together border-t mt-3 pt-1.5 flex justify-between text-[8px] text-slate-400 num-font">
      <span>GalaxyTeck v4.1.2 · Gemini 3.8 Flash · Page 1/1</span>
      <span>Copyright © 2024-2026 GALAXY VIETNAM CO., LTD.</span>
    </div>
  </div>

  <div v-show="showAiModal" class="no-print fixed inset-0 bg-black/60 backdrop-blur-sm flex items-center justify-center p-4 z-50">
    <div class="bg-white rounded-2xl shadow-2xl max-w-5xl w-full p-6 max-h-[92vh] flex flex-col border-2" style="border-color: #A855F7;">
      <div class="flex justify-between items-center border-b-2 pb-3 mb-3" style="border-color: #E9D5FF;">
        <div class="flex items-center gap-2">
          <span class="text-base font-bold" style="color: #7C3AED;">🤖 AI 智能导入（Gemini 3.8 Flash）</span>
          <span v-if="aiStatus.available" class="badge-navy">✅ {{ aiStatus.model }}</span>
          <span v-else class="badge-gold">⚠️ 未配置</span>
        </div>
        <button @click="closeAiModal" class="text-slate-500 hover:text-slate-800 font-bold text-lg">✕</button>
      </div>

      <div v-if="!aiStatus.available" class="p-4 rounded-xl text-sm mb-4" style="background: #FEF3C7; border: 2px solid #F59E0B; color: #92400E;">
        <strong>⚠️ AI 功能未启用</strong>
        <p class="mt-1 text-xs">请设置环境变量 GEMINI_API_KEY 后重启服务：</p>
        <code class="block mt-2 p-2 bg-white rounded text-[11px] num-font">[Environment]::SetEnvironmentVariable("GEMINI_API_KEY", "AIzaSy...", "User")</code>
      </div>

      <div class="rounded-xl p-3.5 mb-3 border-2" style="background: #FAF5FF; border-color: #E9D5FF;">
        <div class="font-bold text-xs mb-2" style="color: #7C3AED;">🎯 导入模式</div>
        <div class="flex gap-3 text-xs">
          <label class="flex items-center gap-2 cursor-pointer"><input type="radio" v-model="aiImportMode" value="library" class="accent-purple-600">📚 仅导入物料库</label>
          <label class="flex items-center gap-2 cursor-pointer"><input type="radio" v-model="aiImportMode" value="customer" class="accent-purple-600">👤 导入 + 客户报价单</label>
        </div>
        <div v-if="aiImportMode==='customer'" class="pt-2 mt-2 border-t" style="border-color: #E9D5FF;">
          <select v-model.number="aiCustomerId" class="fin-input">
            <option :value="0">-- 选择客户 --</option>
            <option v-for="c in customerList" :key="c.id" :value="c.id">{{ c.company_name }}</option>
          </select>
        </div>
      </div>

      <div class="border-2 border-dashed rounded-xl p-5 mb-3 text-center" style="background: #F8FAFC; border-color: #CBD5E1;">
        <input type="file" ref="aiFileInput" @change="handleAiFileSelect" accept="image/*,.pdf" class="hidden">
        <div v-if="!aiFile">
          <div class="text-4xl mb-2">📷</div>
          <p class="text-xs text-slate-600 mb-3">上传报价单图片或 PDF（支持手写、扫描、多语言）</p>
          <button @click="$refs.aiFileInput.click()" :disabled="!aiStatus.available" class="btn-ai px-5 py-2.5 text-xs disabled:opacity-50">
            📁 选择文件
          </button>
        </div>
        <div v-else>
          <p class="text-xs font-bold" style="color: #059669;">✅ {{ aiFile.name }} ({{ (aiFile.size/1024).toFixed(1) }} KB)</p>
          <div class="flex justify-center gap-2 mt-3">
            <button @click="parseWithAI" :disabled="aiParsing || !aiStatus.available" class="btn-ai px-5 py-2 text-xs disabled:opacity-50">
              {{ aiParsing ? '⏳ AI 识别中...' : '🚀 开始 AI 识别' }}
            </button>
            <button @click="aiFile=null; aiProducts=[]" class="px-4 py-2 bg-slate-200 hover:bg-slate-300 rounded-xl text-xs font-bold">重选</button>
          </div>
          <p v-if="aiParsing" class="text-[10px] text-slate-500 mt-2">💡 若遇服务器繁忙，系统会自动重试（最多 8 次），请耐心等待...</p>
        </div>
      </div>

      <div v-if="aiProducts.length > 0" class="flex-1 flex flex-col min-h-0">
        <div class="flex justify-between items-center rounded-xl p-2.5 mb-2 text-xs border-2" style="background: #ECFDF5; border-color: #10B981;">
          <div class="flex items-center gap-3 font-bold" style="color: #047857;">
            <span>✅ 识别 {{ aiProducts.length }} 项</span>
            <span v-if="aiMeta.confidence" class="badge-navy">置信度 {{ (aiMeta.confidence*100).toFixed(0) }}%</span>
            <span v-if="aiMeta.detected_customer" class="badge-gold">👤 {{ aiMeta.detected_customer }}</span>
            <span v-if="aiMeta.from_cache" class="badge-navy">💾 缓存命中</span>
          </div>
          <button @click="toggleAllAIItems" class="text-[11px] px-2.5 py-1 bg-white border-2 rounded font-bold" style="color: #047857; border-color: #10B981;">
            {{ allAIConfirmed ? '取消全选' : '全选' }}
          </button>
        </div>
        <div class="overflow-y-auto flex-1 border-2 rounded-xl" style="border-color: #E2E8F0;">
          <table class="w-full text-xs">
            <thead class="fin-thead sticky top-0">
              <tr>
                <th class="p-2 w-10"><input type="checkbox" :checked="allAIConfirmed" @change="toggleAllAIItems" class="accent-blue-600"></th>
                <th class="p-2 text-left">型号</th><th class="p-2 text-left">规格</th><th class="p-2">单位</th>
                <th class="p-2 text-center">数量</th><th class="p-2 text-right">单价</th><th class="p-2 text-left">备注</th>
              </tr>
            </thead>
            <tbody>
              <tr v-for="(p, i) in aiProducts" :key="i" :class="['border-b', p.confirmed ? '' : 'opacity-50']" style="border-color: #E2E8F0;">
                <td class="p-2 text-center"><input type="checkbox" v-model="p.confirmed" class="accent-blue-600"></td>
                <td class="p-2"><input v-model="p.product_code" class="fin-input text-[11px] font-semibold"></td>
                <td class="p-2"><input v-model="p.product_spec" class="fin-input num-font text-[11px]"></td>
                <td class="p-2"><input v-model="p.unit" class="fin-input text-center text-[11px]"></td>
                <td class="p-2"><input type="number" v-model.number="p.quantity" class="fin-input num-font text-center text-[11px]"></td>
                <td class="p-2"><input type="number" v-model.number="p.unit_price" class="fin-input num-font text-right font-bold text-[11px]"></td>
                <td class="p-2"><input v-model="p.note" class="fin-input text-[11px]"></td>
              </tr>
            </tbody>
          </table>
        </div>
        <div class="flex justify-end gap-2 pt-3 border-t mt-3">
          <button @click="closeAiModal" class="px-4 py-2 bg-slate-100 hover:bg-slate-200 rounded-xl text-xs font-bold">取消</button>
          <button @click="submitAIImport" :disabled="aiImporting" class="btn-ai px-5 py-2 text-xs disabled:opacity-50">
            {{ aiImporting ? '导入中...' : '✅ 确认导入' }}
          </button>
        </div>
      </div>
    </div>
  </div>

  <div v-show="showExcelModal" class="no-print fixed inset-0 bg-black/60 backdrop-blur-sm flex items-center justify-center p-4 z-50">
    <div class="bg-white rounded-2xl shadow-2xl max-w-5xl w-full p-6 max-h-[92vh] flex flex-col border-2" style="border-color: #2563EB;">
      <div class="flex justify-between items-center border-b-2 pb-3 mb-3" style="border-color: #DBEAFE;">
        <span class="text-base font-bold" style="color: #1E3A8A;">📥 Excel 智能导入</span>
        <button @click="showExcelModal=false" class="text-slate-500 hover:text-slate-800 font-bold text-lg">✕</button>
      </div>
      <div class="rounded-xl p-3.5 mb-3 border-2" style="background: #EFF6FF; border-color: #BFDBFE;">
        <div class="flex gap-3 text-xs">
          <label class="flex items-center gap-2 cursor-pointer"><input type="radio" v-model="excelImportMode" value="library" class="accent-blue-600">📚 仅导入物料库</label>
          <label class="flex items-center gap-2 cursor-pointer"><input type="radio" v-model="excelImportMode" value="customer" class="accent-blue-600">👤 导入 + 客户报价单</label>
        </div>
        <div v-if="excelImportMode==='customer'" class="pt-2 mt-2 border-t" style="border-color: #BFDBFE;">
          <select v-model.number="excelCustomerId" class="fin-input">
            <option :value="0">-- 选择客户 --</option>
            <option v-for="c in customerList" :key="c.id" :value="c.id">{{ c.company_name }}</option>
          </select>
        </div>
      </div>
      <div class="border-2 border-dashed rounded-xl p-5 mb-3 text-center" style="background: #F8FAFC; border-color: #CBD5E1;">
        <input type="file" ref="excelFileInput" @change="handleExcelFileSelect" accept=".xlsx,.xls,.csv" class="hidden">
        <div v-if="!excelFile">
          <div class="text-4xl mb-2">📊</div>
          <button @click="$refs.excelFileInput.click()" class="btn-primary px-5 py-2.5 text-xs">📁 选择 Excel 文件</button>
        </div>
        <div v-else>
          <p class="text-xs font-bold" style="color: #059669;">✅ {{ excelFile.name }}</p>
          <button @click="parseExcelFile" :disabled="excelParsing" class="btn-primary mt-2 px-5 py-2 text-xs disabled:opacity-50">
            {{ excelParsing ? '解析中...' : '🔍 解析文件' }}
          </button>
        </div>
      </div>
      <div v-if="excelProducts.length > 0" class="flex-1 flex flex-col min-h-0">
        <div class="rounded-xl p-2.5 mb-2 text-xs font-bold border-2" style="background: #ECFDF5; border-color: #10B981; color: #047857;">
          ✅ 识别 {{ excelProducts.length }} 项
        </div>
        <div class="overflow-y-auto flex-1 border-2 rounded-xl" style="border-color: #E2E8F0;">
          <table class="w-full text-xs">
            <thead class="fin-thead sticky top-0">
              <tr>
                <th class="p-2 w-10"><input type="checkbox" :checked="allExcelConfirmed" @change="toggleAllExcelItems" class="accent-blue-600"></th>
                <th class="p-2 text-left">型号</th><th class="p-2 text-left">规格</th><th class="p-2">单位</th>
                <th class="p-2 text-center">数量</th><th class="p-2 text-right">单价</th><th class="p-2 text-left">备注</th>
              </tr>
            </thead>
            <tbody>
              <tr v-for="(p, i) in excelProducts" :key="i" :class="['border-b', p.confirmed ? '' : 'opacity-50']" style="border-color: #E2E8F0;">
                <td class="p-2 text-center"><input type="checkbox" v-model="p.confirmed" class="accent-blue-600"></td>
                <td class="p-2"><input v-model="p.product_code" class="fin-input text-[11px]"></td>
                <td class="p-2"><input v-model="p.product_spec" class="fin-input text-[11px]"></td>
                <td class="p-2"><input v-model="p.unit" class="fin-input text-center text-[11px]"></td>
                <td class="p-2"><input type="number" v-model.number="p.quantity" class="fin-input num-font text-center text-[11px]"></td>
                <td class="p-2"><input type="number" v-model.number="p.unit_price" class="fin-input num-font text-right text-[11px]"></td>
                <td class="p-2"><input v-model="p.note" class="fin-input text-[11px]"></td>
              </tr>
            </tbody>
          </table>
        </div>
        <div class="flex justify-end gap-2 pt-3 border-t mt-3">
          <button @click="showExcelModal=false" class="px-4 py-2 bg-slate-100 hover:bg-slate-200 rounded-xl text-xs font-bold">取消</button>
          <button @click="submitExcelImport" :disabled="excelImporting" class="btn-primary px-5 py-2 text-xs disabled:opacity-50">
            {{ excelImporting ? '导入中...' : '✅ 确认导入' }}
          </button>
        </div>
      </div>
    </div>
  </div>

  <div v-show="showCustomerModal" class="no-print fixed inset-0 bg-black/60 backdrop-blur-sm flex items-center justify-center p-4 z-50">
    <div class="bg-white rounded-2xl shadow-2xl max-w-5xl w-full p-6 max-h-[90vh] flex flex-col border-2" style="border-color: #10B981;">
      <div class="flex justify-between items-center border-b-2 pb-3 mb-3" style="border-color: #D1FAE5;">
        <span class="text-base font-bold" style="color: #047857;">👥 CRM 客户管理（共 {{ customerList.length }} 笔）</span>
        <button @click="closeCustomerModal" class="text-slate-500 hover:text-slate-800 font-bold text-lg">✕</button>
      </div>
      <div v-if="viewingCustomerQuotesFor" class="rounded-xl p-3.5 mb-3 space-y-2 border-2" style="background: #EFF6FF; border-color: #BFDBFE;">
        <div class="flex justify-between items-center">
          <span class="font-bold text-sm" style="color: #1E3A8A;">【{{ viewingCustomerQuotesFor.company_name }}】的历史订单</span>
          <button @click="viewingCustomerQuotesFor=null" class="text-xs font-bold" style="color: #1E3A8A;">返回 ✕</button>
        </div>
        <div class="overflow-y-auto max-h-56 border-2 rounded-lg bg-white" style="border-color: #E2E8F0;">
          <table class="w-full text-xs">
            <thead class="fin-thead sticky top-0">
              <tr><th class="p-2.5">单号</th><th class="p-2.5">阶梯</th><th class="p-2.5 text-right">含税总额</th><th class="p-2.5 text-center">操作</th></tr>
            </thead>
            <tbody>
              <tr v-for="q in customerQuotesList" :key="q.id" class="border-b hover:bg-blue-50/50">
                <td class="p-2.5 num-font font-bold" style="color: #1E3A8A;">{{ q.quote_no }}</td>
                <td class="p-2.5">{{ q.customer_tier }}</td>
                <td class="p-2.5 text-right num-font font-bold">{{ formatCurrency(q.final_total_amount) }} {{ q.quote_currency }}</td>
                <td class="p-2.5 text-center">
                  <button @click="loadQuoteFromCRM(q.id)" class="btn-primary px-3 py-1 text-xs">调出维护</button>
                </td>
              </tr>
              <tr v-if="customerQuotesList.length===0"><td colspan="4" class="p-6 text-center text-slate-400 text-xs">无订单记录</td></tr>
            </tbody>
          </table>
        </div>
      </div>
      <div v-else class="p-4 rounded-xl mb-3 text-xs space-y-2.5 border-2" style="background: #F8FAFC; border-color: #E2E8F0;">
        <div class="font-bold flex justify-between" style="color: #1E3A8A;">
          <span>{{ editingCustomer.id ? '✏ 编辑客户' : '➕ 新增客户' }}</span>
          <button v-if="editingCustomer.id" @click="resetCustomerForm" class="text-[11px] underline">切换为新增</button>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-3 gap-2.5">
          <input v-model="editingCustomer.company_name" placeholder="* 客户公司全称" class="fin-input font-bold">
          <input v-model="editingCustomer.tax_code" placeholder="税号" class="fin-input num-font">
          <input v-model="editingCustomer.contact_person" placeholder="联系人" class="fin-input">
          <input v-model="editingCustomer.phone" placeholder="电话" class="fin-input num-font">
          <input v-model="editingCustomer.email" placeholder="Email" class="fin-input">
          <select v-model="editingCustomer.default_tier" class="fin-input">
            <option>大型经销商</option><option>中型经销商</option>
            <option>中型安装商</option><option>小型安装商</option>
          </select>
          <input v-model="editingCustomer.address" placeholder="地址" class="md:col-span-2 fin-input">
          <div class="flex gap-2">
            <button @click="submitSaveCustomer" class="flex-1 btn-primary py-2 text-xs">{{ editingCustomer.id ? '保存修改' : '立即新增' }}</button>
          </div>
        </div>
      </div>
      <div class="overflow-y-auto flex-1 border-2 rounded-xl" style="border-color: #E2E8F0;">
        <table class="w-full text-xs">
          <thead class="fin-thead sticky top-0">
            <tr><th class="p-3 text-left">公司名称</th><th class="p-3">税号</th><th class="p-3">联系人</th><th class="p-3">阶梯</th><th class="p-3 text-center">订单</th><th class="p-3 text-center w-40">操作</th></tr>
          </thead>
          <tbody>
            <tr v-for="c in customerList" :key="c.id" class="border-b hover:bg-blue-50/30" style="border-color: #F1F5F9;">
              <td class="p-3 font-bold text-slate-800">{{ c.company_name }}</td>
              <td class="p-3 num-font text-[11px] text-slate-500">{{ c.tax_code || '-' }}</td>
              <td class="p-3">{{ c.contact_person }} <span class="text-slate-400">({{ c.phone || '-' }})</span></td>
              <td class="p-3"><span class="badge-navy">{{ c.default_tier }}</span></td>
              <td class="p-3 text-center">
                <button @click="openCustomerQuotes(c)" :class="['px-2.5 py-1 rounded-lg text-xs font-bold border-2 transition', c.quote_count>0 ? 'bg-sky-100 text-sky-800 border-sky-300 hover:bg-sky-200' : 'bg-slate-100 text-slate-500 border-slate-200']">
                  📜 ({{ c.quote_count || 0 }})
                </button>
              </td>
              <td class="p-3 text-center space-x-1">
                <button @click="selectCustomer(c); closeCustomerModal()" class="btn-success px-2.5 py-1 text-[11px]">带入</button>
                <button @click="editCustomer(c)" class="px-2 py-1 bg-slate-100 hover:bg-slate-200 rounded-lg text-[11px] font-bold">改</button>
                <button @click="deleteCustomer(c.id)" class="px-2 py-1 bg-rose-50 hover:bg-rose-100 rounded-lg text-[11px] font-bold" style="color: #DC2626;">删</button>
              </td>
            </tr>
          </tbody>
        </table>
      </div>
      <div class="flex justify-end pt-3 border-t mt-3">
        <button @click="closeCustomerModal" class="px-5 py-1.5 bg-slate-100 hover:bg-slate-200 rounded-xl text-xs font-bold">关闭</button>
      </div>
    </div>
  </div>

  <div v-show="showSearchModal" class="no-print fixed inset-0 bg-black/60 backdrop-blur-sm flex items-center justify-center p-4 z-50">
    <div class="bg-white rounded-2xl shadow-2xl max-w-4xl w-full p-6 max-h-[85vh] flex flex-col border-2" style="border-color: #2563EB;">
      <div class="flex justify-between items-center border-b-2 pb-3 mb-3" style="border-color: #DBEAFE;">
        <span class="text-base font-bold" style="color: #1E3A8A;">🔍 产品库 (共 {{ products.length }} 款)</span>
        <button @click="showSearchModal=false" class="text-slate-500 hover:text-slate-800 font-bold text-lg">✕</button>
      </div>
      <div class="flex gap-2.5 mb-3">
        <input v-model="modalKeyword" ref="modalInput" placeholder="输入型号、规格、品牌..." class="flex-1 fin-input">
        <select v-model="selectedCategoryFilter" class="fin-input max-w-[180px]">
          <option value="">全部类别</option>
          <option v-for="cat in availableCategories" :key="cat" :value="cat">{{ cat }}</option>
        </select>
      </div>
      <div class="overflow-y-auto flex-1 border-2 rounded-xl" style="border-color: #E2E8F0;">
        <table class="w-full text-xs">
          <thead class="fin-thead sticky top-0">
            <tr><th class="p-2.5 text-left">类别</th><th class="p-2.5 text-left">型号</th><th class="p-2.5 text-left">规格</th><th class="p-2.5 text-right">底价</th><th class="p-2.5 text-center">操作</th></tr>
          </thead>
          <tbody>
            <tr v-for="p in searchResults" :key="p.id" class="border-b hover:bg-blue-50/30" style="border-color: #F1F5F9;">
              <td class="p-2.5 text-[11px] text-slate-500">{{ p.category }}</td>
              <td class="p-2.5 font-bold text-slate-800">{{ p.model }}</td>
              <td class="p-2.5 num-font text-[10px] text-slate-500">{{ p.spec }}</td>
              <td class="p-2.5 text-right num-font font-bold" style="color: #1E3A8A;">{{ formatCurrency(getBasePrice(p)) }}</td>
              <td class="p-2.5 text-center">
                <button @click="addSpecificProduct(p)" class="btn-primary px-3 py-1 text-xs">+ 选入</button>
              </td>
            </tr>
          </tbody>
        </table>
      </div>
      <div class="flex justify-end pt-3 border-t mt-3">
        <button @click="showSearchModal=false" class="px-5 py-1.5 bg-slate-100 hover:bg-slate-200 rounded-xl text-xs font-bold">关闭</button>
      </div>
    </div>
  </div>

  <div v-show="showHistoryModal" class="no-print fixed inset-0 bg-black/60 backdrop-blur-sm flex items-center justify-center p-4 z-50">
    <div class="bg-white rounded-2xl shadow-2xl max-w-4xl w-full p-6 max-h-[80vh] flex flex-col border-2" style="border-color: #64748B;">
      <div class="flex justify-between items-center border-b-2 pb-3 mb-3" style="border-color: #E2E8F0;">
        <span class="text-base font-bold" style="color: #1E3A8A;">📂 历史单据库 ({{ historyList.length }})</span>
        <button @click="showHistoryModal=false" class="text-slate-500 hover:text-slate-800 font-bold text-lg">✕</button>
      </div>
      <div class="overflow-y-auto flex-1 border-2 rounded-xl" style="border-color: #E2E8F0;">
        <table class="w-full text-xs">
          <thead class="fin-thead sticky top-0">
            <tr><th class="p-3">ID</th><th class="p-3 text-left">单号</th><th class="p-3 text-left">客户</th><th class="p-3 text-right">含税总额</th><th class="p-3 text-center">状态</th><th class="p-3 text-center">时间</th><th class="p-3 text-center">操作</th></tr>
          </thead>
          <tbody>
            <tr v-for="q in historyList" :key="q.id" class="border-b hover:bg-blue-50/30" style="border-color: #F1F5F9;">
              <td class="p-3 text-center num-font text-slate-400">{{ q.id }}</td>
              <td class="p-3 num-font font-bold" style="color: #1E3A8A;">{{ q.quote_no }}</td>
              <td class="p-3 font-medium text-slate-700">{{ q.customer_name }}</td>
              <td class="p-3 text-right num-font font-bold" style="color: #1E3A8A;">{{ formatCurrency(q.final_total_amount) }} {{ q.quote_currency }}</td>
              <td class="p-3 text-center"><span class="badge-navy">{{ q.status }}</span></td>
              <td class="p-3 text-center num-font text-[10px] text-slate-400">{{ q.created_at }}</td>
              <td class="p-3 text-center">
                <button @click="loadQuoteDetail(q.id)" class="btn-primary px-3 py-1 text-[11px]">调出维护</button>
              </td>
            </tr>
            <tr v-if="historyList.length===0"><td colspan="7" class="text-center py-8 text-slate-400 text-xs">暂无历史记录</td></tr>
          </tbody>
        </table>
      </div>
      <div class="flex justify-end pt-3 border-t mt-3">
        <button @click="showHistoryModal=false" class="px-5 py-1.5 bg-slate-100 hover:bg-slate-200 rounded-xl text-xs font-bold">关闭</button>
      </div>
    </div>
  </div>

</div>

<script>
const { createApp } = Vue;

createApp({
  data() {
    return {
      currentLang: 'zh',
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
      editingCustomer: { id: null, company_name: '', tax_code: '', contact_person: '', phone: '', email: '', address: '', default_tier: '中型经销商' },

      quickSearchKeyword: '',
      modalKeyword: '',
      selectedCategoryFilter: '',
      historyList: [],

      companyInfo: {
        logo_data: null,
        name_vi: 'CÔNG TY TNHH TẬP ĐOÀN CÔNG NGHỆ GALAXY VIỆT NAM',
        name_cn: '银河科技集团（越南）有限公司',
        tax_code: '2301296587',
        contact_info: 'Điện thoại: 0325.609.218 | Email: galaxyteck6886@gmail.com',
        address: 'Số 36 Yết Kiêu, Phường Kinh Bắc, Tỉnh Bắc Ninh, Việt Nam',
        bank_info: 'STK (Techcombank): 314431 - Ngân hàng TMCP Kỹ thương Việt Nam'
      },

      layoutConfig: {
        title_text: 'BẢNG BÁO GIÁ 商业报价单',
        greeting_text: 'Chúng tôi xin gửi đến Quý khách hàng bảng báo giá như sau / 我司现向贵司呈报以下报价表:'
      },

      form: {
        quote_no: '', layout_template_id: 1, customer_id: 0,
        customer_name: '', contact_person: '', phone_email: '',
        sales_rep: 'Thanh Bình 清平', customer_tier: '中型经销商',
        delivery_scenario: '越南本地仓', trade_term: 'DAP',
        target_currency: 'VND', vat_rate_pct: 10.0,
        valid_days: 3, markup_pct: 0.0, lang_mode: 'vi_zh', template_style: 'standard'
      },

      termsList: [
        'Báo giá có hiệu lực trong vòng 3 ngày / 本报价有效期为 3 天。',
        'Thời gian giao hàng: 3-7 ngày / 交货期：3-7 天内。',
        'Giá chưa bao gồm vận chuyển và lắp đặt / 单价不含运输及安装。',
        'Địa điểm giao hàng: Kho khách hàng chỉ định / 交货地点：客户指定仓库。',
        'Thanh toán: cọc 50% khi xác nhận, 50% trước khi giao hàng / 付款：确认预付 50%，发货前付清 50%。'
      ],
      tableRows: []
    };
  },
  computed: {
    t() {
      const dict = {
        vi: { workspace: 'Bàn làm việc Báo giá Nhanh', customer_view: '👔 Khách hàng', finance_view: '📊 Tài chính' },
        zh: { workspace: '智慧快速报价工作台', customer_view: '👔 客户版', finance_view: '📊 财务版' },
        en: { workspace: 'Smart Quotation Workspace', customer_view: '👔 Customer', finance_view: '📊 Finance' }
      };
      return (k) => (dict[this.currentLang] || dict.zh)[k] || k;
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
      if (lang === 'en') { this.layoutConfig.title_text = 'QUOTATION 报价单'; }
      else { this.layoutConfig.title_text = 'BẢNG BÁO GIÁ 商业报价单'; }
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
          if (this.aiProducts.length === 0) alert('⚠️ AI 未识别到产品，请检查图片清晰度');
        } else {
          alert('❌ AI 识别失败：' + (d.detail || '未知错误'));
        }
      } catch(e) { alert('❌ 请求失败：' + e); } finally { this.aiParsing = false; }
    },
    toggleAllAIItems() { const nv = !this.allAIConfirmed; this.aiProducts.forEach(p => p.confirmed = nv); },
    async submitAIImport() {
      const items = this.aiProducts.filter(p => p.confirmed);
      if (!items.length) return alert('⚠️ 请至少勾选一项');
      if (this.aiImportMode === 'customer' && !this.aiCustomerId) return alert('⚠️ 请选择客户');
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
          if (d.quotation_no) msg += `\\n\\n📄 已生成报价单：${d.quotation_no}`;
          alert('🎉 ' + msg);
          await Promise.all([this.loadProducts(), this.loadCustomers(), this.loadHistory()]);
          if (d.quotation_id) await this.loadQuoteDetail(d.quotation_id);
          this.closeAiModal();
        } else alert('❌ 导入失败：' + (d.detail || '未知'));
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
        } else alert('❌ 解析失败：' + (d.detail || '未知'));
      } finally { this.excelParsing = false; }
    },
    toggleAllExcelItems() { const nv = !this.allExcelConfirmed; this.excelProducts.forEach(p => p.confirmed = nv); },
    async submitExcelImport() {
      const items = this.excelProducts.filter(p => p.confirmed);
      if (!items.length) return alert('⚠️ 请至少勾选一项');
      if (this.excelImportMode === 'customer' && !this.excelCustomerId) return alert('⚠️ 请选择客户');
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
        } else alert('❌ 导入失败：' + (d.detail || '未知'));
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
    resetCustomerForm() { this.editingCustomer = { id: null, company_name: '', tax_code: '', contact_person: '', phone: '', email: '', address: '', default_tier: '中型经销商' }; },
    async submitSaveCustomer() {
      if (!this.editingCustomer.company_name) return alert('请填写公司全称');
      const r = await fetch('/api/customers/save', { method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify(this.editingCustomer) });
      const d = await r.json();
      if (d.status === 'success') { alert('✅ 保存成功'); this.resetCustomerForm(); this.loadCustomers(); }
    },
    async deleteCustomer(id) {
      if (!confirm('确定删除该客户？')) return;
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
      if (p.category && p.category.includes('组件')) unit = '块 / Tấm';
      else if (p.category && p.category.includes('逆变器')) unit = '套 / Bộ';
      this.tableRows.push({
        product_id: p.id, name: p.model, spec: p.spec || p.model, unit,
        quantity: 1, base_price: this.getBasePrice(p), custom_markup: null, note: ''
      });
    },
    addCustomRow() {
      this.tableRows.push({ product_id: 0, name: '自定义产品', spec: '', unit: '个 / Cái', quantity: 1, base_price: 0, custom_markup: null, note: '' });
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
    addTermRow() { this.termsList.push('新增条款内容...'); },
    removeTermRow(i) { this.termsList.splice(i, 1); },
    moveTermUp(i) { if (i > 0) { [this.termsList[i-1], this.termsList[i]] = [this.termsList[i], this.termsList[i-1]]; } },
    moveTermDown(i) { if (i < this.termsList.length-1) { [this.termsList[i+1], this.termsList[i]] = [this.termsList[i], this.termsList[i+1]]; } },

    async handleNewQuote() {
      if ((this.tableRows.length || this.form.customer_name) && !confirm('⚠️ 确认要新建？当前内容将重置')) return;
      this.currentQuoteId = null;
      this.form.customer_id = 0; this.form.customer_name = '';
      this.form.contact_person = ''; this.form.phone_email = '';
      this.form.markup_pct = 0; this.form.quote_no = '';
      this.initDate(); this.tableRows = [];
      if (this.products.length) this.addSpecificProduct(this.products[0]);
      await this.fetchNextQuoteNo();
    },

    async saveToServer(forceNew = false) {
      if (!this.tableRows.length) return alert('⚠️ 请先添加明细');
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
        } else alert('❌ ' + (d.detail || '保存失败'));
      } catch(e) { alert('❌ 请求失败：' + e); }
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
        alert(`✅ 已调出报价单：${h.quote_no}`);
      } catch(e) { alert('❌ 加载失败：' + e); }
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

# ================= 启动 =================
if __name__ == '__main__':
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)