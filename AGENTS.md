# 党政机关公文自动排版引擎（AutoHotkey v2）

基于 GB/T 9704-2012 的 Word/WPS 公文自动排版工具。仓库不是常规工程：**没有构建、包管理、测试、lint 或 CI**，只有按版本号命名的单文件 AHK 脚本和文档。

## 先读这些

- `6.8.12.ahk`：当前最新版，唯一应修改的脚本；约 3900 行，全部逻辑在这一个文件里。
- `README.md`：面向用户的行为说明和逐版本更新日志。
- `程序整体说明.md`：整体架构、函数分区、数据存储和已知坑，换模型接手时先看它。
- `备忘录.txt`：用户的 10 条协作约定（中文交流、先规划、改版本号等）。

## 修改流程（用户的硬性约定）

1. 先只做规划，**得到用户明确同意后才写代码**。
2. 每次改程序都要**提升版本号并另存为新文件**（如 `6.8.12.ahk` → `6.8.13.ahk`），不要直接改旧版本。
3. 新版本文件内必须同步更新：文件头 `V6.8.x`、`global APP_VERSION := "6.8.x"`、文件头「修改记录」中新增一条；再更新 `README.md` 的「当前版本」。
4. 代码必须写**详细中文注释**，尤其解释非直观的 COM/Word 操作和边界条件。
5. 改完要更新 `程序整体说明.md`，并**自动运行程序**供用户测试。
6. 只有确实修改了代码时，回复的最后一行才写 `已完成`；纯讨论/文档不需要。

## 运行与验证（AHK、Git 都不在 PATH）

```powershell
# 启动 GUI
& 'C:\AutoHotkey\v2\AutoHotkey64.exe' '6.8.12.ahk'

# 静默批量：文件或文件夹作为参数；文件夹只扫第一层，~$ 锁文件跳过
& 'C:\AutoHotkey\v2\AutoHotkey64.exe' '6.8.12.ahk' '示例.docx' '某文件夹'

# Git（仓库根目录就是当前目录）
& 'W:\git\bin\git.exe' status --short
& 'W:\git\bin\git.exe' log --oneline
```

没有测试套件。验证方式：把 `6.8.12.ahk` 和一份 `.docx` 复制到临时目录，静默跑两次，检查日志是否按“一文件一行”追加；再启动 GUI 让用户测试。

## 日志（v6.8.11 起）

- 文件名固定为程序目录下的 `公文排版执行日志.txt`，**不带日期时间，追加写入，不覆盖**。
- 每个文件处理完立即写一行，成功、失败、成功但有提示都写；全部成功也写。
- 行格式：`时间<Tab>结果<Tab>文件路径<Tab>说明`；说明中的换行/Tab 会替换为空格，保证严格一行。没有说明时省略第四列。
- 写入函数是 `AppendExecutionLog()`；`SanitizeLogField()` 负责清理字段。旧的覆盖式 `WriteLogFile()` 已删除。
- 排版失败先查日志。最常见原因是目标文档正被 Word/WPS 打开（“文件被占用”）。

## 设置与数据（全部在注册表）

- 设置只在 `HKEY_CURRENT_USER\Software\DocFormatTool`，**没有 `config.ini`，也不要再引入**。
- 子项：`Margins`、`PageNumber`、`FontScheme`、`Options`、`Profile`、`Profiles\<方案名>`、`FontSchemes\<方案名>`。
- 当前方案参数同时写在顶层键，排版逻辑读顶层键；`Profiles` 保存整套方案。内置只读方案名为「党政机关公文格式」。
- 导出设置：GUI「设置 → 导出当前所有设置」生成程序目录下的 `公文排版设置.reg`；换机双击导入后重启。
- 备份默认在 `A_MyDocuments\公文自动排版文件备份`，文件名 `原名_yyyyMMdd_HHmmss.ext`；任务列表右键可恢复最近备份。

## 架构要点

- 入口在 `6.8.12.ahk` 末尾：有命令行参数 → `AppendTasksToQueue()` + `RunSilentProcessing()`；无参数 → `ShowGUI()`。
- 两条排版流水线：GUI 列表用 `RunGUIProcessing()`，静默/右键用 `RunSilentProcessing()`；两者最终都调用 `FormatWordDocument()`。
- `FormatWordDocument()` 是核心引擎：校验占用 → 备份 → 打开文档 → 自动编号转文本 → 页边距 → 清理空段 → 表格保护快照 → 识别大标题/副标题块/主送/正文/附件/落款/日期 → 页码 → 保存/PDF → 恢复表格。
- `GetWordApplication()` 按 **WPS（Kwps / WPS）优先、Microsoft Word 最后** 的顺序尝试 COM，并记录后台 PID 到 `ManagedWordPIDs`。
- GUI 主窗口在 `ShowGUI()`；方案参数编辑在 `OpenProfileEditor()`；字体/版式偏离国标时用红色星号标记。
- 多实例通过 `%TEMP%\DocFormatTool_TaskQueue.txt` 和 `DocFormatTool_Queue.lock` 串行化，锁超过 30 秒视为陈旧。

## 编辑注意事项

- 源文件是 **UTF-8（无 BOM）+ CRLF**。改动时保持原编码和换行；用 PowerShell `Get-Content` 看中文可能乱码，先设 `[Console]::OutputEncoding=[Text.Encoding]::UTF8`，或直接用 read 工具。
- 只在 `.doc`/`.docx` 上工作，跳过 `~$` 开头的 Word 锁文件。
- 文档被占用时不要强行改：提示用户关闭 Word/WPS，或让用户点「结束后台排版进程」（只杀程序自己启动的 Word）。
- 表格保护选项 `OptProtectTable` 勾选时会先快照、排版后恢复表格格式；取消勾选才按方案统一表格字体字号。
- 改右键菜单相关代码时注意：未编译运行 `"<AutoHotkey路径>" "<脚本路径>" "%1"`，编译后运行 EXE；路径修复逻辑在 `CheckAndRepairContextMenuPath()`。

## Git 状态陷阱

- 旧版本大多已从工作区删除，只保留最新脚本（当前为 `6.8.10.ahk`、`6.8.11.ahk`、`6.8.12.ahk`）；`git status` 会显示 15 个旧版本文件为已删除（`D`）。这是用户有意清理，**不要 `git add -A` 或提交这些删除**；需要旧版本时用 `git show HEAD:6.8.9.ahk` 查看。
- 没有配置远程仓库。每个版本按用户要求单独提交为一个 `.ahk` 文件。
- `.gitignore` 已排除 `.reg`、`.doc/.docx`、`.pdf`、日志、`.exe`、备份目录；不要提交这些生成物。
