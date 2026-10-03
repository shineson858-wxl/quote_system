# 贡献指南 / Contributing Guide

感谢你愿意帮助优化 GALAXYTECK 报价系统！本项目面向社区开放，欢迎提交改进。中英文均可，也欢迎越南语。

Thank you for helping to improve GALAXYTECK! This project is open to the community — contributions in Chinese, English or Vietnamese are all welcome.

---

## 欢迎哪些帮助 / What Help Is Welcome

- **代码优化 / Code optimization**：性能、内存、并发、数据库查询优化
- **重构 / Refactoring**：合并多个 main 版本、消除重复代码、拆分模块
- **安全加固 / Security**：配置管理、密钥处理、输入校验、防注入
- **测试 / Testing**：单元测试、接口测试、回归测试
- **功能增强 / Features**：新品类、新模板、新报表、多语言完善
- **文档 / Documentation**：注释补全、使用说明、FAQ
- **Bug 修复 / Bug fixes**：请在提交前说明复现步骤

## 流程 / Workflow

1. **Fork** 本仓库到你的账号 / Fork this repository to your account
2. 创建分支：`git checkout -b feature/your-improvement` / Create a branch
3. 提交修改：`git commit`，提交信息建议使用中英双语或英文 / Commit with clear messages
4. 推送分支并创建 **Pull Request** / Push the branch and open a Pull Request
5. 维护者 review 后合并 / A maintainer will review and merge

## 代码规范 / Code Style

- 遵循 **PEP 8** 基本规范 / Follow PEP 8 basics
- 注释建议中英双语（本项目注释为中英混合风格） / Comments in Chinese + English (matching the project style)
- **禁止在代码中硬编码任何 API 密钥、密码等敏感信息**，统一走环境变量或配置文件 / Never hardcode secrets — always use environment variables or config files
- 保留现有文件顶部的版权声明，不要删除 / Keep the existing copyright headers
- 修改前后请自行运行验证，不要提交无法运行的代码 / Verify your changes run before submitting
- 涉及数据库变更时，请同步说明需要执行的 SQL / Document any DB schema changes

## 优先事项 / Priority Areas

当前最希望得到帮助的方向（详见 README「已知问题」）：

1. 整合 5 个 main 版本为单一代码库 / Consolidate the 5 main versions into one codebase
2. 将硬编码配置（数据库连接、汇率）迁移到环境变量 / Move hardcoded config (DB, rates) to env vars
3. 补充自动化测试 / Add automated tests
4. 前后端拆分 / Split frontend from backend
5. 完善日志与错误处理 / Improve logging & error handling

## 提交信息示例 / Commit Message Examples

```
feat: 新增越南语报价模板 / add Vietnamese quotation template
fix: 修复价格计算精度问题 / fix price rounding issue
refactor: 抽取数据库连接为公共模块 / extract shared DB module
test: 为报价接口添加单元测试 / add unit tests for quote API
```

## 联系方式 / Contact

项目由 GALAXY VIỆT NAM 科技集团维护。如有疑问，可在 Issue 中提出，或通过 Pull Request 直接沟通。

Maintained by GALAXY VIỆT NAM Technology Group. Questions are welcome via Issues or Pull Requests.
