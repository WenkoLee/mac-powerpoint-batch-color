# 本地评审说明

## 项目名称

- 中文：**一键改色**
- 英文：**One-Click Color for Mac PowerPoint**
- GitHub 仓库名：`mac-powerpoint-batch-color`
- 版本：**0.1.2**

## 已备齐

- 独立可读源码、离线构建和本地打包流程。
- MIT 许可证、第三方原文及来源声明。
- 中文普通用户说明、英文项目简介、开发说明。
- 自制示例与真实 PowerPoint 截图；无客户演示文稿。
- 上一版 99 项应用内功能断言；本版默认范围及面板检查，详细边界见验收记录。
- 普通用户安装包、完整源码包和 SHA256 校验文件。

从 [项目首页](README.md)、[使用指南](docs/USER_GUIDE.md)、[验收记录](docs/acceptance/ACCEPTANCE.md) 开始检查。

## 兼容性边界

实体 Command 快捷键、部分文本选择、文本撤销/重做、受组织策略管理的宏设置，以及其他 Mac/Office 组合尚未全部验证。图表原有的自动继承和主题关联不属于完整格式回退保证；详见 [兼容性](docs/COMPATIBILITY.md)。

## 维护者确认与发布

维护者已于 2026-10-13 确认名称、许可证、用户说明、截图及测试版范围，并授权发布到 GitHub。公开仓库：[https://github.com/WenkoLee/mac-powerpoint-batch-color](https://github.com/WenkoLee/mac-powerpoint-batch-color)；插件下载入口见项目首页。后续修改仍需更新版本、重新构建，并按影响范围验收。

维护者于 2026-10-14 要求统一更名“一键改色”，取消名称中的测试版标记，并要求后续修改同步本机与 GitHub。
