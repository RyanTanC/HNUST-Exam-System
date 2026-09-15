# AGENTS.md

本文件是给后续维护者、AI 代理和新手开发者看的项目协作说明。进入本仓库后，先读这里，再读 `README.md` 和 `项目说明.txt`。

本项目是真实文件系统中的项目，请把数据安全放在第一位。除非用户明确要求并确认，不要删除文件、清空目录、批量移动资源或重写题库。

## 1. 项目是什么

`HNUST仿真平台` 是一个面向湖南科技大学学生的 Python 程序设计考试练习桌面应用，用来模拟学校机房考试系统的练习流程。

核心能力：

- 从 `题库/` 目录读取 Excel 试卷。
- 支持单选、判断、填空、程序填空、程序改错、程序设计等题型。
- 支持自动判分、即时反馈、成绩页、答题进度记录。
- 程序题可调用本机 Python IDLE 打开对应 `.py` 文件。
- 支持程序文件备份和重做恢复。
- 支持深色/浅色主题、字体大小设置、考试倒计时、题目导航。
- 支持 GitHub Release 更新检查。
- 使用 PyInstaller 打包成 Windows 单文件 exe。

主要技术栈：

- Python
- PySide6
- PySide6-Fluent-Widgets（qfluentwidgets，设置页等界面组件）
- pandas
- openpyxl
- requests
- pytest
- PyInstaller

## 2. 重要目录与文件

根目录：

- `main.py`：程序入口。只负责设置 PyInstaller 下 Qt 插件路径，然后调用 `hnust_exam.app.run()`。
- `requirements.txt`：运行和测试依赖。
- `HNUST仿真平台.spec`：推荐的 PyInstaller 打包配置。
- `README.md`：面向用户的项目介绍、下载说明和更新日志。
- `项目说明.txt`：面向新手的完整操作手册，包含文件说明、常见修改方法、运行测试和打包步骤。
- `icon.ico`：软件图标，打包时会带入。
- `pay.jpg`：README 使用的支持作者图片。
- `题库/`：Excel 试卷和程序题源文件目录，是核心资源。
- `tests/`：自动测试，当前主要覆盖判分和配置读写。
- `docs/TELEMETRY.md`：遥测相关说明。
- `backup.py`：旧版或备份代码。除非用户明确要求，不要修改、删除或以它为新功能入口。

核心包 `hnust_exam/`：

- `hnust_exam/app.py`：应用总开关。创建 `QApplication`、加载配置、设置主题和全局 QSS、打开主窗口、记录崩溃日志。
- `hnust_exam/models/`：数据模型。
  - `question.py`：单题结构 `Question`。
  - `exam.py`：试卷加载、答案状态、题目标记、当前题目位置、判分入口。
  - `result.py`：判分结果结构。
- `hnust_exam/services/`：服务层。
  - `grader.py`：判分逻辑，包含填空题、程序题和整卷判分。
  - `config_manager.py`：配置和进度读写。
  - `backup_manager.py`：程序题文件备份与恢复。
  - `python_env.py`：查找系统 Python，调用 IDLE 打开程序题文件。
  - `update_checker.py`：GitHub Release 更新检查。
  - `telemetry.py`：遥测相关逻辑。
- `hnust_exam/views/`：PySide6 界面。
  - `main_window.py`：主窗口和页面切换。
  - `welcome_page.py`：欢迎页。
  - `select_page.py`：选卷页。
  - `exam_page.py`：考试页，包含计时、导航、交卷、程序题打开等交互。
  - `question_widget.py`：题目展示和作答区域。
  - `nav_panel.py`：右侧题目导航。
  - `result_page.py`：成绩页。
  - `dialogs/`：设置、交卷确认、Python 环境、更新提示等对话框。
  - `widgets/`：自定义小组件。
- `hnust_exam/utils/`：通用工具。
  - `constants.py`：版本号、仓库名、配置路径、题型顺序、考试时长、Excel 必要列。
  - `helpers.py`：资源路径、答案标准化、代码标准化、版本比较、日志目录。
  - `theme.py`：浅色/深色主题颜色和字体缩放。
  - `ui_helpers.py`：统一风格的消息框等 UI 辅助函数。
- `hnust_exam/viewmodels/`：目前预留给未来扩展。

## 3. 运行、测试、打包

优先使用项目根目录下的 `.venv`。

运行程序：

```powershell
.venv/Scripts/python.exe main.py
```

运行全部测试：

```powershell
.venv/Scripts/python.exe -m pytest tests/ -v
```

只运行判分测试：

```powershell
.venv/Scripts/python.exe -m pytest tests/test_grader.py -v
```

只运行填空题测试类：

```powershell
.venv/Scripts/python.exe -m pytest tests/test_grader.py::TestCheckFillIn -v
```

语法检查：

```powershell
.venv/Scripts/python.exe -m compileall -q hnust_exam tests main.py
```

> **测试不许碰用户真实数据**。`tests/conftest.py` 里有一个 autouse 的
> `isolate_user_data` fixture，会把 `~/.hnust_exam` 下的配置、进度、题库路径
> 全部重定向到临时目录。新增测试直接依赖它就行，不用自己再 patch 一遍；
> 如果新加了从 `_CONFIG_DIR` / `QUESTION_BANK_DIR` 派生的路径常量，
> 记得同步补进那个 fixture。
>
> 历史上踩过两次坑：一次是测试把假数据写进了真实 `manifest.json`，一次是
> `settings_dialog._get_current_version()` 里的 `importlib.reload(constants)`
> 把运行期改过的路径常量打回初值，导致测试把配置写到了真实目录。前者靠这个
> fixture 兜住，后者已改成只扫源文件、不再 reload。

推荐打包方式（优先使用 .spec 文件）：

```powershell
.venv/Scripts/pyinstaller.exe HNUST仿真平台.spec
```

> .spec 文件采用精简策略：PySide6 只引入实际使用的 5 个子模块（QtCore/QtGui/QtWidgets/QtSvg/QtSvgWidgets），
> 避免 collect_all 拖入 WebEngine（~195MB）等无用组件。numpy/pandas/requests 等由 PyInstaller
> 的 import graph 自动发现。仅显式补充 `openpyxl`（hidden import）和 `certifi` 的 `cacert.pem`（SSL 证书）。
>
> QtSvg / QtSvgWidgets / QtXml 是 qfluentwidgets 带来的（`SettingCard` 用了 `QSvgWidget`，
> `common/icon.py` 用了 QtXml）。qfluentwidgets 的样式和图标都内嵌在 `.py` 里，包内没有任何
> 数据文件，所以**不需要** `collect_data_files("qfluentwidgets")`。
> 它自带的 `qfluentwidgets.multimedia` 子包会拉 QtMultimedia，但顶层 `__init__` 并不导入它，
> 正常使用不会被打进包里——实测确认过，别顺手加 collect_all。

备用打包命令（不使用 .spec 文件时）：

```powershell
.venv/Scripts/pyinstaller.exe --onefile --windowed --icon=icon.ico --name "HNUST仿真平台" --add-data "题库;题库" --add-data "icon.ico;." --add-data "hnust_exam/resources;hnust_exam/resources" --hidden-import PySide6.QtCore --hidden-import PySide6.QtGui --hidden-import PySide6.QtWidgets --hidden-import openpyxl --collect-data certifi --runtime-tmpdir "%TEMP%\HNUST_simulation_temp" main.py
```

> `--collect-data certifi` 是必需的：SSL 根证书 `cacert.pem` 不会被 PyInstaller 自动收集，
> 缺失会导致 HTTPS 请求（更新检查、遥测等）运行时失败。其余依赖通过 import graph 自动发现。

打包产物：

```text
dist/HNUST仿真平台.exe
```

打包前确认：

- `main.py` 能正常运行。
- `题库/` 存在且包含试卷。
- `icon.ico` 在项目根目录。
- 版本号、README 更新日志、GitHub Release tag 保持一致。
- 相关测试通过。

## 4. 常见修改位置

改考试时长：

- 修改 `hnust_exam/utils/constants.py` 里的 `EXAM_TIME_SECONDS`。
- 例如 30 分钟是 `1800`，60 分钟是 `3600`，90 分钟是 `5400`。

改版本号：

- 修改 `hnust_exam/utils/constants.py` 里的 `CURRENT_VERSION`。
- 同时检查 `README.md` 更新日志和 GitHub Release tag，避免更新检查误判。

改 GitHub 更新检测仓库：

- 修改 `hnust_exam/utils/constants.py` 里的 `GITHUB_USERNAME` 和 `GITHUB_REPO_NAME`。

改判分逻辑：

- 填空题：改 `hnust_exam/services/grader.py` 的 `check_fill_in()`。
- 程序题：改 `check_program_answer()`，尤其是 `seq.ratio() >= 0.9` 的相似度阈值。
- 判断题答案归一化：改 `hnust_exam/utils/helpers.py` 的 `normalize_answer()`。

改 Excel 题库读取：

- 必要列：改 `hnust_exam/utils/constants.py` 的 `REQUIRED_COLUMNS`。
- 加载和分组：改 `hnust_exam/models/exam.py` 的 `load_from_excel()`。
- 题目字段：改 `hnust_exam/models/question.py`。

改主题颜色：

- 修改 `hnust_exam/utils/theme.py`。
- `_LIGHT` 和 `_DARK` 中相同语义的颜色键要同步考虑。

改页面 UI：

- 欢迎页：`hnust_exam/views/welcome_page.py`。
- 选卷页：`hnust_exam/views/select_page.py`。
- 考试页：`hnust_exam/views/exam_page.py`。
- 题目展示：`hnust_exam/views/question_widget.py`。
- 导航栏：`hnust_exam/views/nav_panel.py`。
- 成绩页：`hnust_exam/views/result_page.py`。
- 对话框：`hnust_exam/views/dialogs/`。

改设置项：

- 默认值放在 `hnust_exam/services/config_manager.py` 的 `DEFAULTS`。
- UI 控件通常放在 `hnust_exam/views/dialogs/settings_dialog.py`。
- 读取设置的位置要就近查找，不要引入全局状态。

> 设置页用的是 qfluentwidgets。三个坑：
>
> 1. `SettingCardGroup` 内部的 `ExpandLayout` 用 `widget.height()`（**当前**高度）
>    累加分组高度，而不是 `sizeHint()`。标准 `SettingCard` 靠 `setFixedHeight`
>    兜底，所以自定义卡片（`_SectionCard`）必须在内容填完后调用
>    `sync_height()`，否则会被压扁成一条空白。内容显隐/字号变化后也要再调一次。
> 2. `app.py` 的全局 QSS 只给 `QMainWindow` 上底色，`QDialog` 拿不到；而且它会
>    样式化 `QPushButton`/`QLineEdit` 等基础类。设置对话框靠 `_apply_local_style()`
>    自己下发底色和中性化规则，切换主题预览时要重新调一次。
> 3. **主题色要对齐**。qfluentwidgets 自带主色是青色 `#009faa`，和 `Theme.PRIMARY`
>    （浅色 `#0078d7` / 深色 `#4da6ff`）不是一回事，不对齐的话设置页的开关、滑块、
>    进度条、主按钮会和软件其它部分两个色。`_sync_fluent_theme()` 负责这件事，
>    主题变化时要重新调。
>
>    已知差异：浅色下渲染出来和 `Theme.PRIMARY` **完全一致**；深色下
>    `ThemeColor.PRIMARY.color()` 会做一次 `s *= 0.84, v = 1` 的变换（库为了让
>    深色背景上的强调色不刺眼，是刻意设计），实际画出来略淡（`#4da6ff` → `#69b4ff`）。
>    色相一致，**不要反向补偿**——那等于把库的内部实现细节焊死在代码里。
>
>    `setThemeColor()` 必须带 `save=False`（这也是它的默认值，别改）：qfluentwidgets
>    的配置文件默认落在**相对路径** `config/config.json`，也就是进程当前目录，
>    带 `save=True` 会在用户的工作目录里拉一坨垃圾。`tests/test_settings_dialog.py`
>    里有对应用例锁住这几点。

## 5. 数据与文件安全

最高优先级：保护用户数据和题库资源。

绝对禁止：

- 递归删除目录或内容，例如 `Remove-Item -Recurse`、`rm -rf`、`rmdir /s`。
- 通配符批量删除，例如 `del *.tmp`、`Remove-Item *.bak`、`rm *.log`。
- 删除整个目录，即使目录为空也不要自动执行。
- 用循环、管道、脚本批量删除多个文件。
- 使用静默或强制删除参数，例如 `/q`、`-f`。

唯一允许的删除方式：

- 一次只删除一个明确路径的单个普通文件。
- 路径必须是完整、硬编码、无通配符的文件路径。
- 如果不能确定目标是不是单个普通文件，停止并让用户手动检查。

遇到批量删除、清空文件夹、删除目录、按模式删除的请求时：

- 立即拒绝自动执行。
- 明确说明触发了文件安全规则。
- 建议用户打开资源管理器手动处理。
- 不提供替代自动化删除脚本。

特别谨慎的区域：

- `题库/`：包含试卷和程序题文件，不能随意改名、删除或批量格式化。
- `hnust_exam/services/backup_manager.py`：涉及程序题文件备份和恢复，修改前必须明确数据影响。
- 用户目录 `.hnust_exam/`：保存配置和答题进度。
- `dist/`、`build/`：打包产物和构建缓存，不主动清理。
- `.venv/`、`.idea/`、`.claude/`：本地环境或工具配置，不主动改动。

## 6. 编码协作原则

### 6.1 先想清楚再改

- 不要假设需求；不确定就说明假设或询问。
- 如果有多个解释，先列出差异，不要悄悄选一个。
- 如果更简单的方案能解决问题，优先简单方案。
- 每次修改前明确成功标准，例如“新增测试覆盖无效输入并通过”。

### 6.2 简单优先

- 不加用户没有要求的功能。
- 不为一次性逻辑抽象复杂框架。
- 不提前做“以后可能会用”的配置项。
- 如果改动明显可以更小，就缩小范围。

### 6.3 手术式修改

- 只改和任务直接相关的文件。
- 不顺手重构无关代码。
- 不顺手格式化整仓库。
- 匹配现有代码风格，即使你个人会写成另一种风格。
- 发现无关死代码可以报告，但不要主动删除。

### 6.4 目标驱动验证

常见任务应转成可验证目标：

- “改判分” -> 添加或更新 `tests/test_grader.py`，再让测试通过。
- “改配置” -> 添加或更新 `tests/test_config.py`，验证读写。
- “改 UI” -> 至少运行程序手动检查对应页面。
- “改打包” -> 运行 PyInstaller 或说明未运行的原因。

修改后优先运行最小相关测试；无法运行时，在回复中说明原因。

## 7. 发布前检查

发布前至少检查：

- `CURRENT_VERSION` 是否正确。
- `README.md` 更新日志是否同步。
- GitHub Release tag 是否和版本号一致。
- `requirements.txt` 是否包含新增依赖。
- `HNUST仿真平台.spec` 是否包含必要资源。
- `题库/` 是否能被打包后读取。
- `python -m pytest tests/ -v` 是否通过。
- `dist/HNUST仿真平台.exe` 是否能在干净环境启动。

## 8. 题库 Excel 常见问题

新增或修改试卷 Excel 时，注意以下几点：

**`程序文件` 列必须使用相对文件名**，如 `Prog00001.py`，文件应位于 `题库/试题文件夹/` 下。绝对路径（如 `D:\Exam\07000143\CK\23`）在其他机器上不存在，会导致"找不到程序文件"错误。

**文件名拼写检查**：确保 `.py` 后缀完整（`Prog2.3.py` 而非 `Prog2.3py`）。

**程序文件命名规则**：

- `Prog00001.py` ~ `Prog00023.py`：程序填空/改错（`PY程序填空.xlsx`、`PY程序改错.xlsx`）
- `Prog00024.py` ~ `Prog00033.py`：程序设计模拟题（`PY 程序设计模拟题.xlsx`）
- `Prog1.1.py` ~ `Prog5.10.py`：实验模拟题 1~5（`PY程序设计实验-模拟题N.xlsx`）
- `Prog1.1.1.py` ~ `Prog1.5.3.py`：程序设计模拟题 1~5（`PY程序设计模拟题N.xlsx`）

**排查试卷问题的步骤**：

1. 用 openpyxl 读取 Excel，检查 `程序文件` 列的值。
2. 确认 `题库/试题文件夹/` 下存在对应文件。
3. 如有绝对路径，需将源文件复制到 `试题文件夹/` 并更新 Excel 引用。

**Excel 文件被占用时的处理**：写入时若报 `PermissionError`，先保存到临时文件，让用户关闭占用程序后再替换。

**程序文件编码**：`试题文件夹/` 中的 `.py` 文件使用 cp936 编码（中文注释），不要用 UTF-8 重新保存，否则中文注释会乱码。

## 9. 给 AI 代理的额外提醒

- 本项目是中文用户项目，文档和用户可见文本优先使用中文。
- 不要把 README 的用户宣传语和 AGENTS 的开发协作说明混在一起；AGENTS 更偏维护指南。
- 运行命令前确认当前目录是项目根目录。
- 修改涉及 Windows 路径、中文文件名、PyInstaller 参数时，保留现有写法，避免无谓跨平台重写。
- 对打包命令中的 `--add-data "题库;题库"` 保持 Windows 分号格式。
- 如果要新增依赖，同步更新 `requirements.txt` 并说明用途。
- 不要自动提交、推送或创建 PR，除非用户明确要求。

