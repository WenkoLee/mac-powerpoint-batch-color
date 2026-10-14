# 独立源码与构建

## 目录

```text
src/                 可读 VBA 源码、界面控件清单、功能区 XML
scripts/             构建、审计与发布包脚本
templates/           空白宿主种子和构建用加载项容器
vendor/pyopenvba/    已验证的 MIT 授权依赖快照
licenses/            第三方许可证原文
tests/               静态测试和 PowerPoint 内验收代码
examples/            自制示例，不含客户资料
docs/                用户说明、截图和验收记录
release/             当前安装包及校验值
build/               临时构建与验收文件，Git 忽略
```

## 环境

需要 Python 3.10+，仅使用 Python 标准库和已随项目提供的 pyOpenVBA 6.1.2 快照。构建不要求安装 PowerPoint，不联网、不读取旧开发目录，不需要安装其他 Python 包。

如果 macOS 自带的 `python3` 低于 3.10，请显式使用已安装的较新 Python 路径。

## 构建

```sh
python3 scripts/build.py
```

构建过程：

1. 读取 `src/` 的可读 UTF-8 VBA 源码。
2. 将中文字符串转换为 VBA 的 `ChrW$` 表达式，避免宿主代码页差异导致乱码。
3. 按 `form-controls.json` 创建 99 个界面控件，附加面板代码。
4. 使用清理后的空白宿主项目生成 VBA 工程。
5. 将工程和功能区 XML 放入构建用 `.ppam` 容器，输出安装包。
6. 写出 SHA256 校验值和源码哈希清单。

表单身份和 ZIP 时间戳固定。同一源码、依赖快照及 Python 环境应生成字节相同的安装包；测试会检查两次构建是否一致。

## 修改源码

- `ColorReplaceInPlace.bas`：范围解析、对象遍历、颜色匹配、替换及撤销。
- `ColorPickerForm.bas`：中文无模式面板、对象/选页事件、预览、文本编辑与 macOS 纯文本剪贴板读取。
- `form-controls.json`：表单的结构。运行时位置和样式由面板代码设置。
- `ribbon.xml`：PowerPoint 顶部入口。

直接修改当前源码后重建，无需运行旧版生成脚本。版本号位于 `VERSION`；更新版本时也应同步功能区提示、说明和变更记录。

## 检查

```sh
python3 -m unittest discover -s tests -p 'test_*.py' -v
python3 scripts/audit.py
```

静态检查验证构建可复现、VBA 模块与表单、功能区、ZIP 关系、第三方依赖快照及公开文件清单。它不能确认 Mac PowerPoint 运行表现。

## PowerPoint 内验收

```sh
python3 scripts/build.py --with-qa
```

这会额外生成 `build/AcceptanceHarness.pptm`。只在该测试稿中运行 `RunAcceptanceTests`。它以 `templates/qa-fixture.pptx` 的自制原生图表为起点，用 PowerPoint 原生对象创建其余测试页，并把断言结果写入第 1 页，期间会出现三次插件撤销的结果提示。Mac 不提供 Shapes.AddChart/AddChart2，因此验收不调用该接口。测试完成后在 PowerPoint 中正常保存测试稿，再读取结果。

测试稿附加验收模块，改色引擎与面板源码和当前安装包一致。最终安装包不包含 `AcceptanceRunner`，也没有测试按钮。

完成后还需用真实 `.ppam` 检查加载、界面、页面选择、纯文本粘贴、关闭重开及宏保护恢复后的状态。实际通过范围写入验收记录。

## 构建材料与来源

`powerpoint-base.pptm` 是清理后的空白 PowerPoint 宿主项目，只保留一个空模块和必要类型库引用。它不含业务页面、照片或客户数据。

`addin-container.zip` 是构建用资源容器，不是可运行加载项；沿用 MIT 授权的 ProDeck 容器结构，已去掉原 VBA 工程和功能区代码。构建时使用本项目源码补齐。授权声明见 `licenses/ProDeck-MIT.txt`。

`vendor/SOURCE.json` 记录随项目提供的 pyOpenVBA 快照及逐文件哈希。无需将原仓库的测试文档、个人工作目录或其他插件安装包带入本项目。

`templates/qa-fixture.pptx` 是项目自制的 4 页空白测试种子，只有第 2 页含虚构柱形图，数据为 A/B/C=3/5/4，不含业务数据。此模板无需在每次构建时重新生成。

## 本地打包

```sh
python3 scripts/package.py
```

先执行资料审计，再把源码包和普通用户安装包写入 `dist/`，同时生成 SHA256 校验文件。此步骤不连接 GitHub。`build/`、`dist/`、缓存和 Office 临时文件不进入源码包；GitHub 发布与本地构建分开执行；发布前检查文档、验收范围、第三方授权和资料审计。

若需重新生成图表测试种子，使用 `scripts/make_qa_fixture.py`；这个可选工具需要自行安装 `python-pptx`（本次生成环境为 1.0.2），日常构建和验收使用仓库中已提供的模板，不需要该依赖。
