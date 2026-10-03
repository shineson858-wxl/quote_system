# GALAXYTECK 智慧快速报价与客户关系管理系统 / Smart Quotation & CRM System

**企业标准网络版 / Enterprise Standard Web Edition**

GALAXYTECK 是面向越南光伏、通风与制冷设备行业的智能报价与客户关系管理系统，提供基于 Google Gemini 的 AI 智能报价、客户管理、产品与阶梯价格管理、报价单模板管理等能力，帮助业务人员快速生成专业的中越双语报价单。

GALAXYTECK is an AI-assisted quotation & CRM system built for the solar PV, ventilation and cooling equipment industry in Vietnam. Powered by Google Gemini, it helps sales teams generate professional bilingual (Chinese–Vietnamese) quotations quickly, and manage customers, products, tiered pricing and quotation templates in one place.

> 🌟 欢迎高手参与优化 / Expert contributions are warmly welcome!
> 贡献指南请见 / See [CONTRIBUTING.md](CONTRIBUTING.md)

---

## 功能特性 / Features

- **AI 智能报价**：接入 Google Gemini，自动解析需求并生成报价内容 / AI-assisted quoting powered by Google Gemini
- **客户关系管理（CRM）**：客户档案、报价历史与跟进记录 / Customer profiles, quotation history & follow-ups
- **产品与价格体系**：多品类产品库、阶梯价格矩阵（大型经销商 / 中型经销商 / 中型安装商 / 小型安装商）、多区域与贸易条款（FOB / DAP） / Product library with tiered pricing (dealer / installer tiers), multi-region & trade terms (FOB / DAP)
- **汇率管理**：内置汇率表，支持 VND / CNY / USD 换算 / Exchange-rate management (VND / CNY / USD)
- **报价单模板**：可视化模板与布局管理，支持中越双语 / Quotation template & layout management (Chinese–Vietnamese)
- **Excel 批量导入**：从报价 Excel 一键导入产品与价格到 MySQL / Bulk import of products & prices from Excel into MySQL
- **覆盖品类**：光伏组件、逆变器、储能电池、工业大吊扇（HVLS）、工业冷风机及配件 / Categories: PV modules, inverters, batteries, HVLS ceiling fans, industrial evaporative coolers & parts

## 技术栈 / Tech Stack

| 层 / Layer | 技术 / Technology |
|---|---|
| 后端 / Backend | Python 3 · FastAPI · Uvicorn |
| AI 能力 / AI | Google Gemini (google-generativeai) |
| 数据库 / Database | MySQL (XAMPP / phpMyAdmin) · pymysql |
| 数据处理 / Data | pandas · openpyxl |
| 前端 / Frontend | 内嵌 HTML + JavaScript（Tailwind 风格样式） / Embedded HTML + JS |

## 项目结构 / File Structure

| 文件 / File | 说明 / Description |
|---|---|
| `main.py` | 主程序 v4.1.2（当前推荐版本）：FastAPI 服务 + Web 界面 / Main app v4.1.2 (recommended): FastAPI server + Web UI |
| `main370.py` | v3.7.0 Excel Smart Import Edition（旧版本 / older version） |
| `main350.py` | v3.5.0 CRM Order Linkage Edition（旧版本 / older version） |
| `main42bug.py` | v4.1.0 调试变体（实验性质，不建议直接用于生产） / Debug variant (experimental, not for production) |
| `mainbug.py` | v4.0.2 调试变体（实验性质，不建议直接用于生产） / Debug variant (experimental, not for production) |
| `import_data.py` | 导入汇率、逆变器/储能、东南亚组件参考价到 MySQL / Import exchange rates, inverters & SEA module prices |
| `import_fans.py` | 导入 15 款工业吊扇产品与阶梯价 / Import HVLS fans & tiered prices |
| `coolers.py` | 导入 NAKO 工业冷风机及配件 / Import NAKO industrial coolers & parts |
| `test_db.py` | 数据库连接自检脚本 / DB connectivity test |

## 环境要求与运行 / Requirements & Run

- **Python** 3.10+，依赖：`pip install fastapi uvicorn pymysql pandas openpyxl google-generativeai`
- **MySQL**：本地 XAMPP，数据库 `solar_quotation_db`（核心表：`tb_product`、`tb_price_matrix`、`tb_exchange_rate`）
- **Gemini API Key**：通过环境变量 `GEMINI_API_KEY` 提供（程序内已按环境变量方式读取，请勿在代码中硬编码密钥）
- **启动**：在项目目录运行 `python main.py`，浏览器访问 `http://localhost:8000`

## 已知问题与欢迎优化的方向 / Known Issues & Help Wanted

以下方向欢迎高手协助优化（详见 CONTRIBUTING.md）：

- **多版本并存**：`main.py` / `main370.py` / `main350.py` / `mainbug.py` / `main42bug.py` 存在大量重复代码，建议合并重构 / Multiple parallel versions with duplicated code — consolidation & refactoring welcome
- **配置硬编码**：数据库连接、汇率等硬编码在脚本中，建议改为配置文件或环境变量 / Hardcoded DB config & rates — suggest config file or env vars
- **缺少自动化测试**：无单元测试与接口测试 / No automated tests
- **前端内嵌**：HTML/JS 内嵌于后端，建议拆分前后端 / Embedded frontend — could be split into a separate frontend
- **可观测性**：日志、错误处理、性能优化可加强 / Logging, error handling & performance tuning

## 如何贡献 / How to Contribute

请阅读 [CONTRIBUTING.md](CONTRIBUTING.md) 了解贡献流程、代码规范与优先事项。

Please read [CONTRIBUTING.md](CONTRIBUTING.md) for the contribution workflow, code style and priorities.

## 版权 / Copyright

版权所有 © 2024-2026 CÔNG TY TNHH TẬP ĐOÀN CÔNG NGHỆ GALAXY VIỆT NAM（越南 GALAXY VIỆT NAM 科技集团）

Copyright © 2024-2026 GALAXY VIỆT NAM Technology Group Co., Ltd. All rights reserved.

## 许可证 / License

本项目采用 **MIT 许可证** 开源，任何人可自由使用、修改与分享，详见 [LICENSE](LICENSE)。

This project is licensed under the **MIT License** — anyone is free to use, modify and share it. See [LICENSE](LICENSE).
