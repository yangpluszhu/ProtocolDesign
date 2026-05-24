# Contributing to ProtocolDesign

感谢你关注 ProtocolDesign。

本项目是一个面向中文医学研究场景的 Windows 桌面工具，当前核心目标是：在**不牺牲研究方案结构完整性和可审阅性**的前提下，提升回顾性队列研究方案生成的质量、稳定性与可用性。

## 贡献方式

欢迎以下类型的贡献：

- Bug 修复
- 界面体验改进
- 帮助文档完善
- 生成质量优化
- 模型兼容性增强
- 打包与发布流程改进

如果你计划进行较大改动，建议先通过 Issue 说明背景、目标和拟议方案，再开始实现。

## 开始之前

### 1. 环境要求

- Windows 10 / 11
- Python 3.11+
- LibreOffice（可选，仅用于生成帮助文档 PDF）

### 2. 安装依赖

```powershell
python -m pip install -r requirements.txt
```

### 3. 从源码运行

```powershell
python src/ProtocolDesign.py
```

## 建议的开发流程

1. Fork 本仓库并创建功能分支。
2. 在本地完成修改。
3. 运行基本检查。
4. 如涉及 UI 或帮助文档，同步更新截图和文档。
5. 提交 Pull Request，并清晰说明变更目的、影响范围和验证方式。

## 提交前检查

项目目前没有完整自动化测试体系，因此请至少完成以下检查：

### 语法检查

```powershell
python -m py_compile src/ProtocolDesign.py src/protocol_renderer.py src/build_help_docx.py src/resource_crypto.py src/encrypt_resources.py
```

### 帮助文档生成

```powershell
python src/build_help_docx.py
```

### 可选：完整打包验证

```powershell
.\build_windows.ps1 -SkipInstall -SkipPdf
```

### 手工验证

请至少人工检查以下内容：

- 主界面可以正常启动
- 方案摘要文件选择正常
- 提供商切换后相关字段联动正常
- 输出路径与输出文件名填写正常
- 帮助文档可正常打开
- 至少完成一次完整生成链路验证（如果你具备可用模型接口）

## 对代码改动的特别说明

### 1. 请尽量不要破坏当前输出契约

当前项目的提示词层、校验层和 DOCX 渲染层是紧密耦合的。修改以下任一部分时，请同步检查相关实现：

- 结构化摘要表字段
- 正文章节顺序
- Kimi 分段生成结构
- 附表约束
- 帮助文档中的使用说明

### 2. 修改 UI 时请同步更新文档

如果你的改动影响以下内容，请一并更新：

- `README.md`
- `docs/ProtocolDesign_帮助文档.md`
- `src/build_help_docx.py`
- `help_assets/` 中的界面截图

### 3. 不要提交敏感信息

请不要提交以下内容：

- 真实 API Key
- 个人账号配置
- 含敏感数据的研究资料
- 本地环境产生的临时日志或缓存文件

## Pull Request 建议内容

建议在 PR 描述中说明：

- 变更背景
- 主要修改点
- 是否影响 UI / 文档 / 打包
- 验证方式
- 截图（如涉及界面改动）

## 文档与讨论

- 使用帮助：[`docs/ProtocolDesign_帮助文档.md`](docs/ProtocolDesign_帮助文档.md)
- 发布说明：[`CHANGELOG.md`](CHANGELOG.md)

## 许可证

向本项目提交代码，即表示你同意你的贡献将在本仓库当前使用的 [`GPL-3.0`](LICENSE) 许可证下发布。
