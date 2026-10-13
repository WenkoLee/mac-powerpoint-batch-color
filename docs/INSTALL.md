# 安装与卸载

当前版本为 **0.1.0-beta.1 测试版**，安装文件未经过数字签名。

从 [GitHub 测试版发布页](https://github.com/WenkoLee/mac-powerpoint-batch-color/releases/tag/v0.1.0-beta.1) 下载以 `_Mac.zip` 结尾的安装包并解压。普通用户不需要构建源码。

## 安装前

需要 Mac 桌面版 Microsoft PowerPoint，并且所在环境允许使用 VBA 加载项。已测试环境和未验证组合见 [兼容性记录](COMPATIBILITY.md)。

把项目 `release/` 中的 `.ppam` 文件放在一个固定位置。安装后不要移动或删除该文件；PowerPoint 通过这个位置加载插件。

## 安装

1. 打开 PowerPoint，在顶部 macOS 菜单选择 **工具 → PowerPoint 加载项…**。
2. 若安装过旧版“批量改色”，先取消旧版勾选，避免出现重复入口。
3. 点击 **添加**，选择 `MacPowerPointBatchColor_0_1_0_beta_1.ppam`。
4. 对确认来源可信的这份本地插件，按 PowerPoint 提示选择启用宏。
5. 加载成功后，顶部应出现 **Mac 批量改色·测试版** 选项卡。
6. 打开自制示例，点击 **打开选色面板** 检查入口和预览。

不同 PowerPoint 版本的提示和菜单名称可能略有差异。

## 宏保护设置

Microsoft 365 for Mac 默认对包含宏的文件显示警告。请只启用你确认来源的插件，并保持日常宏保护设置。参考 [Microsoft 官方宏设置说明](https://support.microsoft.com/en-gb/office/vba/enable-or-disable-macros-in-microsoft-365-for-mac)。

部分 Mac PowerPoint 版本在加载 `.ppam` 时会再次询问是否关闭宏病毒保护。如果你的环境出现这一提示，临时关闭会影响其他文件的宏执行方式。若选择继续，完成安装和测试后应立即在 **PowerPoint → 首选项/偏好设置 → 安全性** 中恢复“禁用宏并通知”或重新开启宏病毒保护。

恢复后重启 PowerPoint，检查加载项是否仍可运行。若插件被策略阻止，请记录版本和具体提示；本项目不提供持续关闭宏保护的安装方案。由组织统一管理的电脑，请遵守其软件与宏策略。

## 更新

1. 结束正在进行的改色操作并关闭面板。
2. 在加载项列表取消旧版勾选，再添加新版文件。
3. 确认新版入口可用后，再移除旧版登记和文件。

## 停用与卸载

- **停用**：在 **工具 → PowerPoint 加载项…** 中取消勾选。
- **卸载**：选择对应加载项，点击 **移除**，再删除你保存的 `.ppam` 文件。

停用或卸载不会改变已经保存到演示文稿中的颜色。插件的临时撤销记录不会跨 PowerPoint 重启保存。
