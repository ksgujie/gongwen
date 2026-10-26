; 声明脚本所需的最低 AutoHotkey 版本为 v2.0
#Requires AutoHotkey v2.0
; 允许多开运行多个实例
#SingleInstance Off
; 关闭未初始化变量的警告提示
#Warn VarUnset, Off

; =====================================================================
; 党政机关公文自动排版引擎 (基于 GB/T 9704-2012) - V6.8.3 逐行注释版
;
; 修改记录：
; v6.6.1 修复页码格式
; v6.6.5 新增统一版本号变量、标题最大字数拦截阈值设置、关于窗口
; v6.6.7 移除"生成修改对比副本"功能，修复附件说明缩进覆盖问题与右键恢复菜单
; v6.6.8 新增 SetParagraphText 安全替换函数，根治文末成文日期段落分裂产生双日期的 Bug；
;        修正文末落款单位字体为正文字体（三号仿宋）
; v6.6.9 移除落款全角空格填充，改用动态计算右缩进字符数实现上下居中（单位长时右缩进 2 字符）
; v6.7.0 附件及正文各级标题（一至四级）序号及半角点"."后方的空格全部自动剔除；
;        附件说明采用基于当前方案正文字号（ptBody）的动态磅值悬挂缩进
; v6.7.1 1. 标题最大字数默认拦截阈值调整为 35 字（自动升级旧版默认 50 字配置）
;        2. 统一一至四级标题判定：末尾带标点且总长未超阈值的独立短段直接识别为标题
;        3. 新增 Word/WPS 自动编号转纯文本与特殊不可见空白清理，防止自动编号标题漏判
; v6.7.6 1. 行距预设改为与 Word 界面一致的"单倍行距"，并同步加固行距解析器
;        2. 补全文件头常量的中文尾注释（行距规则 0/4/5、边距预设 1/2/3 等）
;        3. 导出 config.ini 时补写 OptProtectTable，避免导出后该项回落默认值
;        4. 统一行距与磅值的空格写法
; v6.7.7 1. 页码写入改为安全模式：保留页脚原有内容，仅替换本程序生成的页码块；
;        2. 同时处理主页脚、首页页脚和偶数页页脚，页脚含其他内容时保留原内容并给出提示；
;        3. PDF 导出失败时在文件状态、悬浮通知和执行日志中明确提示，不再静默失败；
;        4. 新版本按版本号命名（6.7.7.ahk），并启用 Git 版本库管理。
; v6.7.8 1. 修复关闭“排版前生成原始备份”后，备份存放目录的“浏览”“打开”按钮被置灰的问题；
;        2. 备份目录的路径、浏览和打开按钮始终可用（仅在批量排版期间临时锁定）；
;        3. 新增任务列表双击文件用系统默认程序打开功能，方便排版前核对原文件；
;        4. “打开”备份目录时若目录不存在则自动创建，避免空目录提示。
; v6.7.9 1. 修复“浏览...”备份目录只能在该目录子树内选择的问题，改用系统文件夹选择框，
;        支持向上一级及任意磁盘导航；
;        2. “排版前生成原始备份”由界面复选框移入主菜单“设置”，避免误点关闭；
;        3. 开始排版前若备份开关处于关闭状态，弹出“是/否”提醒对话框（默认“是”），
;        选择“是”即重新开启备份并继续排版（界面批量与静默排版均生效）；
;        4. 批量排版期间同步禁用主菜单中的备份开关，防止运行中误改。
; v6.8.0 1. 停用 config.ini，全部设置只保存在注册表 HKEY_CURRENT_USER\Software\DocFormatTool；
;        2. 自定义字体方案同样只读写注册表，不再与配置文件双向同步；
;        3. “导出设置”改为导出 .reg 注册表文件，换一台电脑双击导入即可恢复全部设置；
;        4. 状态栏改为显示“设置保存在注册表”；
;        5. 导出文件的默认位置改为程序所在目录（编译后为 EXE 目录，未编译为脚本目录），
;        文件名固定为“公文排版设置.reg”，不再附加日期时间。
; v6.8.1 1. 导出设置成功后的模态对话框改为右下角浮动提示框，不再打断当前操作；
;        2. 浮动提示框按文字内容自动调整高度，较长的文件路径不会被截断。
; v6.8.2 1. 拖入文件夹时先检查是否包含子文件夹；若包含，弹出“是/否”提示框询问
;        是否扫描子文件夹中的文件（默认“是”）；
;        2. 选择“是”时递归添加子文件夹中的 Word 文件，选择“否”时只添加第一层。
; v6.8.3 1. 执行日志改为固定保存在程序所在目录（编译后为 EXE 目录，未编译为脚本目录）；
;        2. 日志固定为一个文件“公文排版执行日志.txt”，文件名不再附加日期时间，每次覆盖写入。
; =====================================================================

; --- 全局版本号定义 ---
global APP_VERSION   := "6.8.3"

; --- 默认国标字体名称预设 ---
global FONT_TITLE    := "方正小标宋简体" ; 国标字体名称预设（参考值，运行时以当前字体方案为准）
global FONT_SUBTITLE := "楷体"          ; 国标字体名称预设（参考值，运行时以当前字体方案为准）
global FONT_AUTHOR   := "楷体"          ; 国标字体名称预设（参考值，运行时以当前字体方案为准）
global FONT_H1       := "黑体"          ; 国标字体名称预设（参考值，运行时以当前字体方案为准）
global FONT_H2       := "楷体"          ; 国标字体名称预设（参考值，运行时以当前字体方案为准）
global FONT_BODY     := "仿宋"          ; 国标字体名称预设（参考值，运行时以当前字体方案为准）

; --- 运行时生效字号与行距变量 ---
global SIZE_TITLE        := 22     ; 字号参考值（运行时以当前字体方案经 ParseFontSize 转换后的磅值为准，二号=22磅）
global SIZE_BODY         := 16     ; 字号参考值（运行时以当前字体方案经 ParseFontSize 转换后的磅值为准，三号=16磅）
global SIZE_TABLE        := 10.5   ; 字号参考值（运行时以当前字体方案经 ParseFontSize 转换后的磅值为准，五号=10.5磅）
global LINE_SPACING_RULE := 4      ; Word行距规则：0=单倍, 4=固定值(磅), 5=多倍行距
global LINE_SPACING_BODY := 28     ; 正文默认行距：固定值 28 磅（单位：磅）
global LINE_SPACING_TITLE:= 35     ; 主标题默认行距：固定值 35 磅（单位：磅，运行时会按主标题字号×1.55 重算）

; --- 注册表存储路径定义（设置不再使用 config.ini） ---
global REG_PATH          := "HKEY_CURRENT_USER\Software\DocFormatTool"
global REG_SCHEMES_PATH  := "HKEY_CURRENT_USER\Software\DocFormatTool\FontSchemes"

; --- 页面版心边距参数（单位：毫米 mm） ---
global MarginType        := 1      ; 边距预设模式：1=党政机关公文格式, 2=WORD默认边距, 3=用户自定义
global MarginTop         := 37     ; 国标天头（上边距）37mm
global MarginBottom      := 35     ; 国标地脚（下边距）35mm
global MarginLeft        := 28     ; 国标订口（左边距）28mm
global MarginRight       := 26     ; 国标翻口（右边距）26mm
global CustomTop         := 0.0    ; 用户自定义上边距缓存
global CustomBottom      := 0.0    ; 用户自定义下边距缓存
global CustomLeft        := 0.0    ; 用户自定义左边距缓存
global CustomRight       := 0.0    ; 用户自定义右边距缓存

; --- 页码格式参数 ---
global PageNumFormat     := "格式数字 — 1 — 居中"
global SkipSinglePageNum := 1      ; 单页文档是否不加页码

; --- 当前选中的字体方案名称与自定义方案字典 ---
global CurrentFontScheme := "党政机关公文格式"
global CustomFontSchemes := Map()

; --- 排版附加功能开关配置 ---
global ExportPDF        := 0       ; 是否同时导出 PDF 副本
global CreateBAK        := 1       ; 排版前是否自动生成原始备份
global OptFixDocNum     := 0       ; 是否自动将发文字号括号规范为六角括号〔〕
global OptPreventOrphan := 0       ; 是否启用标题防孤行
global MaxHeadingLen    := 35      ; 标题最大字符数拦截阈值
global ContextActive    := 0       ; 用户是否主动添加过右键菜单标记

; --- 备份文件存储路径 ---
global DefaultBakDir := A_MyDocuments "\公文自动排版文件备份"
global BakDir        := DefaultBakDir

; --- 右下角无焦点悬浮通知框（Toast）全局控件句柄 ---
global toastGui       := ""
global toastTitleCtrl := ""
global toastMsgCtrl   := ""
global toastIconCtrl  := ""

; --- “打开备份目录”方形按钮所用系统图标句柄（需常驻，不能提前释放） ---
global BakOpenHIcon := 0

; --- 资源管理器右键多选并发队列与进程锁文件路径 ---
global QueueFilePath := A_Temp "\DocFormatTool_TaskQueue.txt"
global QueueLockPath := A_Temp "\DocFormatTool_Queue.lock"

; --- 本程序创建的后台 Word/WPS 进程 PID 追踪表 ---
global ManagedWordPIDs := Map()

; --- 主界面 GUI 与各控件全局引用 ---
global isProgramUpdating := false
global mainGui, uiMarginType, uiTop, uiBottom, uiLeft, uiRight, uiPageFormat, uiSkipSingle
global uiFontScheme, uiBtnManageScheme
global uiPDF, uiFixDocNum, uiProtectTable, uiPreventOrphan, uiSettingsMenu
global uiBakPath, uiBakBtn, uiBakOpen, uiLV, uiBtnRun, uiBtnReset, uiBtnClear, uiBtnKill, uiBtnExport, uiSB
global uiTopUD, uiBottomUD, uiLeftUD, uiRightUD

; 主菜单中“排版前生成原始备份”开关项的名称（多处复用，避免拼写不一致）
global BAK_MENU_LABEL := "排版前生成原始备份"

; =====================================================================
; [系统字体枚举核心函数：从注册表安全读取本机已安装的中文字体]
; =====================================================================
GetSystemChineseFonts() {
    ; 定义默认的常用中文字体列表
    fontList := ["方正小标宋简体", "仿宋", "黑体", "楷体", "宋体"]
    fontMap := Map()
    ; 遍历默认字体列表并写入 Map
    for f in fontList
        fontMap[f] := true

    ; 定义系统注册表中的字体存放路径数组
    regKeys := [
        "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts",
        "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
    ]

    ; 循环读取注册表项中的所有字体键值
    for rootKey in regKeys {
        Loop Reg, rootKey, "V" {
            ; 清理并去除字体名称后面的括号修饰
            cleanName := Trim(RegExReplace(A_LoopRegName, "\s*\([^)]*\)$", ""))
            ; 排除空名称和竖排字体，并且在 Map 中不存在时处理
            if (cleanName != "" && SubStr(cleanName, 1, 1) != "@" && !fontMap.Has(cleanName)) {
                ; 仅匹配包含中文字符的字体名称
                if RegExMatch(cleanName, "[\x{4e00}-\x{9fa5}]") {
                    fontMap[cleanName] := true
                    fontList.Push(cleanName)
                }
            }
        }
    }
    ; 返回最终收集到的字体数组
    return fontList
}

; 初始化全局字体列表
global WpsFontNameList := GetSystemChineseFonts()

; 初始化常用字号列表
global WpsFontSizeList := [
    "初号", "小初", "一号", "小一", "二号", "小二",
    "三号", "小三", "四号", "小四", "五号", "小五",
    "六号", "小六", "七号", "八号",
    "42", "36", "26", "24", "22", "18", "16", "15", "14", "12", "10.5", "9", "7.5"
]

; 初始化行距预设列表
global LineSpacingPresets := [
    "固定值 28 磅", "固定值 30 磅", "单倍行距", "1.5 倍行距", "2 倍行距"
]

; 初始化页码格式预设列表
global PageFormatList := [
    "无",
    "普通数字 1     居左", "普通数字 1     居中", "普通数字 1     居右",
    "格式数字 — 1 — 居左", "格式数字 — 1 — 居中", "格式数字 — 1 — 居右",
    "加总页码 1 / n 居左", "加总页码 1 / n 居中", "加总页码 1 / n 居右"
]

; =====================================================================
; [默认国家标准公文方案参数模型 (GB/T 9704-2012)]
; =====================================================================
GetDefaultSchemeMap() {
    ; 返回包含公文各部件默认字体、字号、加粗属性的对象
    return {
        MTitleFont:   "方正小标宋简体", MTitleSize:   "二号", MTitleBold:   0,
        SubTitleFont: "楷体",           SubTitleSize: "三号", SubTitleBold: 0,
        AuthorFont:   "楷体",           AuthorSize:   "三号", AuthorBold:   0,
        BodyFont:     "仿宋",           BodySize:     "三号", BodyBold:     0,
        H1Font:       "黑体",           H1Size:       "三号", H1Bold:       0,
        H2Font:       "楷体",           H2Size:       "三号", H2Bold:       0,
        H3Font:       "仿宋",           H3Size:       "三号", H3Bold:       1,
        H4Font:       "仿宋",           H4Size:       "三号", H4Bold:       0,
        TableFont:    "仿宋",           TableSize:    "五号", TableBold:    0,
        Spacing:      "固定值 28 磅"
    }
}

; 初始化当前激活的样式对象
global ActiveStyle := GetDefaultSchemeMap()

; =====================================================================
; [配置持久化读写抽象层]
; =====================================================================
GetConfigVal(section, key, defaultVal) {
    global REG_PATH
    ; 6.8.0 起设置只存放在注册表；section 参数仅为兼容既有调用而保留
    try return RegRead(REG_PATH, key)
    catch
        return defaultVal
}

SetConfigVal(section, key, val) {
    global REG_PATH
    ; 只写入注册表，不再创建或同步 config.ini
    try RegWrite(val, "REG_SZ", REG_PATH, key)
}

; =====================================================================
; [自定义字体方案加载、保存与删除函数]
; =====================================================================
LoadFontSchemes() {
    global CustomFontSchemes, REG_SCHEMES_PATH
    CustomFontSchemes := Map()
    def := GetDefaultSchemeMap()

    ; 自定义字体方案只从注册表读取
    Loop Reg, REG_SCHEMES_PATH, "K" {
        sName := A_LoopRegName
        try {
            item := {}
            for k, defaultV in def.OwnProps() {
                val := RegRead(REG_SCHEMES_PATH "\" sName, k)
                item.%k% := InStr(k, "Bold") ? (IsNumber(val) ? Integer(val) : 0) : val
            }
            CustomFontSchemes[sName] := item
        }
    }
}

SaveFontSchemeData(sName, schemeObj) {
    global REG_SCHEMES_PATH
    ; 自定义字体方案只写入注册表
    try {
        subKey := REG_SCHEMES_PATH "\" sName
        for k, v in schemeObj.OwnProps()
            RegWrite(v, "REG_SZ", subKey, k)
    }
}

DeleteFontSchemeData(sName) {
    global REG_SCHEMES_PATH
    try RegDeleteKey(REG_SCHEMES_PATH "\" sName)
}

ExportSettingsToReg() {
    global REG_PATH, uiSB

    ; 先把当前界面参数写入注册表，确保导出的内容与界面一致
    SaveAllCurrentSettings()

    ; 默认保存到程序所在目录：编译后为 EXE 所在目录，未编译时为脚本所在目录
    defaultFile := A_ScriptDir "\公文排版设置.reg"
    regFile := FileSelect("S16", defaultFile, "导出设置（注册表文件）", "注册表文件 (*.reg)")
    if (regFile == "")
        return
    if (SubStr(regFile, -4) != ".reg")
        regFile .= ".reg"

    ; 直接导出本程序的注册表根键，包含全部设置和自定义字体方案
    try
        exitCode := RunWait('reg.exe export "' REG_PATH '" "' regFile '" /y', , "Hide")
    catch {
        ShowFloatingToast("导出失败", "无法调用系统注册表导出工具 reg.exe。", 3000, false)
        return
    }

    if (exitCode != 0 || !FileExist(regFile)) {
        ShowFloatingToast("导出失败", "注册表导出失败，请检查目标路径的写入权限！", 3000, false)
        uiSB.SetText(" 导出失败：无法写入 " regFile)
        return
    }

    uiSB.SetText(" 设置已导出：" regFile)
    ; 用右下角浮动提示替代模态对话框，不打断当前操作；完整路径同时显示在状态栏
    ShowFloatingToast("导出成功", "注册表文件已保存：`n" regFile "`n换电脑双击该文件导入，重新启动即可恢复设置。", 5000, true)
}

ParseFontSize(inputStr, defaultVal := 16.0) {
    str := Trim(inputStr)
    nameMap := Map(
        "初号", 42.0, "小初", 36.0, "一号", 26.0, "小一", 24.0,
        "二号", 22.0, "小二", 18.0, "三号", 16.0, "小三", 15.0,
        "四号", 14.0, "小四", 12.0, "五号", 10.5, "小五", 9.0,
        "六号", 7.5,  "小六", 6.5,  "七号", 5.5,  "八号", 5.0
    )
    if nameMap.Has(str)
        return nameMap[str]
    numOnly := RegExReplace(str, "[^\d\.]", "")
    return (IsNumber(numOnly) && Float(numOnly) > 0) ? Float(numOnly) : defaultVal
}

ParseLineSpacing(inputStr, &ruleOut, &valOut) {
    str := Trim(inputStr)
    numPart := RegExReplace(str, "[^\d\.]", "")
    valNum := (IsNumber(numPart) && Float(numPart) > 0) ? Float(numPart) : 28.0

    if InStr(str, "倍") {
        if (valNum == 1.0 || valNum == 28.0) {
            ; 兼容 "单倍行距"（无数字）与旧版 "1 倍行距"
            ruleOut := 0, valOut := 12
        } else {
            ruleOut := 5, valOut := valNum * 12
        }
    } else {
        ruleOut := 4, valOut := valNum
    }
}

ApplyFontScheme(schemeName) {
    global ActiveStyle, CustomFontSchemes, LINE_SPACING_RULE, LINE_SPACING_BODY, LINE_SPACING_TITLE
    if (schemeName == "党政机关公文格式" || !CustomFontSchemes.Has(schemeName))
        ActiveStyle := GetDefaultSchemeMap()
    else
        ActiveStyle := CustomFontSchemes[schemeName]

    ParseLineSpacing(ActiveStyle.Spacing, &rule, &val)
    LINE_SPACING_RULE := rule
    LINE_SPACING_BODY := val
    titlePt := ParseFontSize(ActiveStyle.MTitleSize, 22.0)
    LINE_SPACING_TITLE:= Round(titlePt * 1.55)
}

; =====================================================================
; [程序主入口]
; =====================================================================
LoadSettings()

; 判断是否为编译后运行且用户主动添加过右键菜单，若路径失效则自愈
if (A_IsCompiled && ContextActive == 1) {
    CheckAndRepairContextMenuPath()
}

; 检查是否存在命令行参数
if (A_Args.Length > 0) {
    AppendTasksToQueue(A_Args)
    if (!TryAcquireMasterLock())
        ExitApp()

    Sleep(350)
    task_files := ReadAndClearQueue()
    ReleaseMasterLock()

    if (task_files.Length > 0)
        RunSilentProcessing(task_files)
    else
        ExitApp()
} else {
    CreateTraySettingsMenu()
    ShowGUI()
}
return

; =====================================================================
; [多文件并发排队队列管理函数]
; =====================================================================
AppendTasksToQueue(args) {
    Loop 20 {
        try {
            f := FileOpen(QueueFilePath, "a", "UTF-8")
            if (f) {
                for arg in args {
                    if (arg != "")
                        f.WriteLine(arg)
                }
                f.Close()
                return true
            }
        } catch {
            Sleep(25)
        }
    }
    return false
}

TryAcquireMasterLock() {
    try {
        if (FileExist(QueueLockPath)) {
            fileTime := FileGetTime(QueueLockPath, "M")
            if (DateDiff(A_Now, fileTime, "Seconds") > 30)
                FileDelete(QueueLockPath)
            else
                return false
        }
        FileAppend(A_Now "`n", QueueLockPath, "UTF-8")
        return true
    } catch {
        return false
    }
}

ReleaseMasterLock() {
    try {
        if (FileExist(QueueLockPath))
            FileDelete(QueueLockPath)
    }
}

ReadAndClearQueue() {
    taskList := []
    uniqueMap := Map()
    if (!FileExist(QueueFilePath))
        return taskList

    Loop 20 {
        try {
            content := FileRead(QueueFilePath, "UTF-8")
            FileDelete(QueueFilePath)
            Loop Parse, content, "`n", "`r" {
                itemPath := Trim(A_LoopField)
                if (itemPath != "" && !uniqueMap.Has(itemPath)) {
                    uniqueMap[itemPath] := true
                    taskList.Push(itemPath)
                }
            }
            break
        } catch {
            Sleep(25)
        }
    }
    return taskList
}

; =====================================================================
; [备份回滚与恢复机制：查找指定文件最近的备份并恢复]
; =====================================================================
RestoreLatestBackup(filePath) {
    global BakDir, DefaultBakDir
    targetDir := (BakDir != "") ? BakDir : DefaultBakDir
    if (!DirExist(targetDir) || !FileExist(filePath)) {
        ShowFloatingToast("恢复失败", "未找到文件或备份目录不存在！", 3000, false)
        return false
    }

    SplitPath(filePath, , , &fileExt, &nameNoExt)
    latestBak := "", latestTime := ""

    Loop Files, targetDir "\" nameNoExt "_*." fileExt {
        if (latestTime == "" || A_LoopFileTimeModified > latestTime) {
            latestTime := A_LoopFileTimeModified
            latestBak := A_LoopFilePath
        }
    }

    if (latestBak != "") {
        try {
            FileCopy(latestBak, filePath, true)
            ShowFloatingToast("恢复成功", "已将文档恢复至最近一次排版前的备份！", 3000, true)
            return true
        } catch {
            ShowFloatingToast("恢复失败", "文件被占用，请先关闭相关 Word 窗口！", 3500, false)
            return false
        }
    } else {
        ShowFloatingToast("恢复失败", "在备份目录中未找到此文件的历史备份！", 3000, false)
        return false
    }
}

; =====================================================================
; [核心界面 (GUI) 构建与事件绑定]
; =====================================================================
ShowGUI() {
    global APP_VERSION
    global mainGui, uiMarginType, uiTop, uiBottom, uiLeft, uiRight, uiPageFormat, uiSkipSingle
    global uiFontScheme, uiBtnManageScheme
    global uiPDF, uiFixDocNum, uiProtectTable, uiPreventOrphan, uiSettingsMenu
    global uiBakPath, uiBakBtn, uiBakOpen, uiLV, uiBtnRun, uiBtnReset, uiBtnClear, uiBtnKill, uiBtnExport, uiSB
    global uiTopUD, uiBottomUD, uiLeftUD, uiRightUD, PageFormatList
    global MarginType, MarginTop, MarginBottom, MarginLeft, MarginRight, PageNumFormat, SkipSinglePageNum
    global CurrentFontScheme, CustomFontSchemes, BAK_MENU_LABEL
    global ExportPDF, CreateBAK, OptFixDocNum, OptProtectTable, OptPreventOrphan, BakDir

    if (IsSet(mainGui) && IsObject(mainGui)) {
        try {
            mainGui.Show()
            return
        }
    }

    mainGui := Gui("+Theme -MaximizeBox", "党政机关公文自动排版引擎 V" APP_VERSION)
    mainGui.OnEvent("Close", GuiClose)
    mainGui.OnEvent("DropFiles", GuiDropFiles)
    mainGui.SetFont("s10", "Microsoft YaHei")
    mainGui.BackColor := "FFFFFF"

    ; --- 顶部菜单栏 ---
    mainMenu := MenuBar()
    settingsMenu := Menu()
    ; 备份开关放在主菜单，避免在界面上被误点关闭；勾选状态即当前是否备份
    settingsMenu.Add(BAK_MENU_LABEL, ToggleBak)
    settingsMenu.Add()
    settingsMenu.Add("添加 Word 文件右键运行命令", (*) => RegisterDocumentContextMenu(true))
    settingsMenu.Add("删除 Word 文件右键运行命令", (*) => RemoveDocumentContextMenu())
    settingsMenu.Add()
    settingsMenu.Add("高级：设置标题最大字数拦截阈值", (*) => OpenHeadingLenSettings())
    settingsMenu.Add()
    settingsMenu.Add("导出当前所有设置 (注册表 .reg)", (*) => ExportSettingsToReg())
    mainMenu.Add("设置(&S)", settingsMenu)
    
    helpMenu := Menu()
    helpMenu.Add("请作者喝一杯咖啡(&A)", (*) => ShowAboutDialog())
    mainMenu.Add("帮助(&H)", helpMenu)
    mainGui.MenuBar := mainMenu

    ; 记录菜单对象，供重置默认值、批量排版锁定和状态同步使用
    uiSettingsMenu := settingsMenu
    UpdateBakMenuCheck()

    ; --- 顶部主横幅 ---
    mainGui.SetFont("s18 bold c0055AA")
    mainGui.Add("Text", "x20 y15 w640 Center BackgroundTrans", "党政机关公文自动排版引擎")
    mainGui.SetFont("s9 norm c666666")
    mainGui.Add("Text", "x20 y45 w640 Center BackgroundTrans", " (符合 GB/T 9704-2012 国标规范)  |  请将 Word 文件拖入下方列表")
    mainGui.SetFont("s10 c000000")

    ; --- 左半区：版心边距设置 ---
    mainGui.Add("GroupBox", "x20 y75 w315 h155", "版心边距设置 (单位：毫米)")
    mainGui.Add("Text", "x32 y105 w60", "边距预设:")
    uiMarginType := mainGui.Add("DropDownList", "x95 y100 w135 Choose" MarginType, ["党政机关公文格式", "WORD默认边距", "用户自定义"])

    mainGui.Add("Text", "x32 y145 w45", "上边距:")
    uiTop := mainGui.Add("Edit", "x78 y140 w60 Center", MarginTop)
    uiTopUD := mainGui.Add("UpDown", "-2 Range0-100", 50)
    mainGui.Add("Text", "x145 y145 w25", "mm")

    mainGui.Add("Text", "x170 y145 w45", "下边距:")
    uiBottom := mainGui.Add("Edit", "x216 y140 w60 Center", MarginBottom)
    uiBottomUD := mainGui.Add("UpDown", "-2 Range0-100", 50)
    mainGui.Add("Text", "x283 y145 w25", "mm")

    mainGui.Add("Text", "x32 y185 w45", "左边距:")
    uiLeft := mainGui.Add("Edit", "x78 y180 w60 Center", MarginLeft)
    uiLeftUD := mainGui.Add("UpDown", "-2 Range0-100", 50)
    mainGui.Add("Text", "x145 y185 w25", "mm")

    mainGui.Add("Text", "x170 y185 w45", "右边距:")
    uiRight := mainGui.Add("Edit", "x216 y180 w60 Center", MarginRight)
    uiRightUD := mainGui.Add("UpDown", "-2 Range0-100", 50)
    mainGui.Add("Text", "x283 y185 w25", "mm")

    ; --- 右半区：页码与字体方案设置 ---
    mainGui.Add("GroupBox", "x345 y75 w315 h155", "格式与字体字号方案")
    mainGui.Add("Text", "x358 y102 w65", "字号方案:")
    schemeList := ["党政机关公文格式"]
    for sName in CustomFontSchemes
        schemeList.Push(sName)

    chooseSchemeIdx := 1
    for idx, item in schemeList {
        if (item == CurrentFontScheme) {
            chooseSchemeIdx := idx
            break
        }
    }
    uiFontScheme := mainGui.Add("DropDownList", "x425 y98 w145 Choose" chooseSchemeIdx, schemeList)
    uiBtnManageScheme := mainGui.Add("Button", "x576 y97 w70 h28", "管理...")
    uiBtnManageScheme.OnEvent("Click", OpenFontSchemeManager)
    uiFontScheme.OnEvent("Change", OnFontSchemeChange)

    mainGui.Add("Text", "x358 y138 w65", "页码格式:")
    chooseFormatIdx := 6
    for idx, item in PageFormatList {
        if (item == PageNumFormat) {
            chooseFormatIdx := idx
            break
        }
    }
    uiPageFormat := mainGui.Add("DropDownList", "x425 y134 w221 Choose" chooseFormatIdx, PageFormatList)
    uiPageFormat.OnEvent("Change", OnPageFormatChange)

    uiSkipSingle := mainGui.Add("Checkbox", "x358 y175 w285 Checked" SkipSinglePageNum, "文件仅一页不添加页码")
    uiSkipSingle.OnEvent("Click", (ctrl, *) => SaveSingleSetting("PageNumber", "SkipSinglePageNum", ctrl.Value))

    uiMarginType.OnEvent("Change", OnMarginTypeChange)
    uiTop.OnEvent("Change", OnMarginEditChange)
    uiBottom.OnEvent("Change", OnMarginEditChange)
    uiLeft.OnEvent("Change", OnMarginEditChange)
    uiRight.OnEvent("Change", OnMarginEditChange)

    uiTopUD.OnEvent("Change", (ctrl, *) => AdjustMarginBySpinner(uiTop, ctrl))
    uiBottomUD.OnEvent("Change", (ctrl, *) => AdjustMarginBySpinner(uiBottom, ctrl))
    uiLeftUD.OnEvent("Change", (ctrl, *) => AdjustMarginBySpinner(uiLeft, ctrl))
    uiRightUD.OnEvent("Change", (ctrl, *) => AdjustMarginBySpinner(uiRight, ctrl))

    uiTop.OnEvent("Focus", EditSelectAll)
    uiBottom.OnEvent("Focus", EditSelectAll)
    uiLeft.OnEvent("Focus", EditSelectAll)
    uiRight.OnEvent("Focus", EditSelectAll)

    ; --- 中部：排版增强与细节控制选项 ---
    mainGui.Add("GroupBox", "x20 y240 w640 h60", "排版增强与细节控制选项")
    uiFixDocNum     := mainGui.Add("Checkbox", "x35 y265 w200 Checked" OptFixDocNum, "规范发文字号六角括号")
    uiProtectTable  := mainGui.Add("Checkbox", "x245 y265 w190 Checked" OptProtectTable, "表格保护与字号适配")
    uiPreventOrphan := mainGui.Add("Checkbox", "x445 y265 w200 Checked" OptPreventOrphan, "标题与下文同页 (防孤行)")

    uiFixDocNum.OnEvent("Click", (ctrl, *) => SaveSingleSetting("Options", "OptFixDocNum", ctrl.Value))
    uiProtectTable.OnEvent("Click", (ctrl, *) => SaveSingleSetting("Options", "OptProtectTable", ctrl.Value))
    uiPreventOrphan.OnEvent("Click", (ctrl, *) => SaveSingleSetting("Options", "OptPreventOrphan", ctrl.Value))

    ; --- 中下部：输出与备份选项 ---
    mainGui.Add("GroupBox", "x20 y310 w640 h105", "输出与备份选项")
    uiPDF  := mainGui.Add("Checkbox", "x35 y335 w200 Checked" ExportPDF, "同时生成 PDF 副本")
    ; 备份开关已移入主菜单“设置”，此处仅作位置提示，避免用户找不到
    mainGui.SetFont("s9 c808080")
    mainGui.Add("Text", "x245 y337 w400", "是否备份请在顶部菜单“设置 → 排版前生成原始备份”中查看或修改")
    mainGui.SetFont("s10 c000000")

    uiPDF.OnEvent("Click", (ctrl, *) => SaveSingleSetting("Options", "ExportPDF", ctrl.Value))

    mainGui.Add("Text", "x35 y373 w90", "备份存放目录:")
    uiBakPath := mainGui.Add("Edit", "x130 y369 w379", BakDir)
    uiBakBtn := mainGui.Add("Button", "x519 y368 w80", "浏览...")
    ; 正方形图标按钮：BS_ICON(0x40) + BM_SETIMAGE 只显示图标，不绘制文字
    uiBakOpen := mainGui.Add("Button", "x609 y368 w31 h31 +0x40", "打开")
    SetButtonStockIcon(uiBakOpen, 4) ; 4 = SIID_FOLDEROPEN（系统标准“打开文件夹”图标）

    ; 备份目录的路径、浏览和打开控件始终可用；是否生成备份由主菜单开关决定
    uiBakPath.Enabled := true
    uiBakBtn.Enabled := true
    uiBakOpen.Enabled := true
    uiBakPath.OnEvent("Change", (ctrl, *) => SaveSingleSetting("Options", "BakDir", ctrl.Value))
    uiBakBtn.OnEvent("Click", SelectBakPath)
    uiBakOpen.OnEvent("Click", OpenBakFolder)

    ; --- 下部：拖拽任务列表区 ---
    mainGui.Add("GroupBox", "x20 y425 w640 h165", "排版任务列表 (拖拽文件至此，双击打开)")
    uiLV := mainGui.Add("ListView", "x30 y450 w620 h130 Grid", ["文件名", "状态", "文件路径"])
    uiLV.ModifyCol(1, 200)
    uiLV.ModifyCol(2, 160)
    uiLV.ModifyCol(3, 230)

    ; 绑定任务列表右键菜单：支持从备份恢复
    lvMenu := Menu()
    lvMenu.Add("从备份恢复此文件", (*) => RestoreSelectedLVItem())
    uiLV.OnEvent("ContextMenu", (ctrl, item, isRightClick, x, y) => (item > 0 ? lvMenu.Show() : ""))
    ; 双击任务列表中的文件，用系统默认程序打开，便于排版前核对原文件
    uiLV.OnEvent("DoubleClick", (*) => OpenSelectedLVFile())

    ; --- 底部：功能操作按钮栏 ---
    uiBtnReset := mainGui.Add("Button", "x20 y600 w80 h35", "默认参数")
    uiBtnReset.OnEvent("Click", ResetDefaults)

    uiBtnClear := mainGui.Add("Button", "x105 y600 w80 h35", "清空列表")
    uiBtnClear.OnEvent("Click", ClearList)

    uiBtnExport := mainGui.Add("Button", "x190 y600 w85 h35", "导出设置")
    uiBtnExport.OnEvent("Click", (*) => ExportSettingsToReg())

    uiBtnKill := mainGui.Add("Text", "x280 y600 w130 h35 Center +0x200 BackgroundFFF9C4 Border c554400", "结束后台排版进程")
    uiBtnKill.SetFont("s10 bold")
    uiBtnKill.OnEvent("Click", KillAllWordProcesses)

    OnMessage(0x0200, CheckButtonHover)

    uiBtnRun := mainGui.Add("Text", "x415 y600 w245 h35 Center +0x200 BackgroundE8F5E9 Border c117711", "保存参数并运行  ▶")
    uiBtnRun.SetFont("s11 bold")
    uiBtnRun.OnEvent("Click", RunGUIProcessing)

    modeText := " [设置保存在注册表]"
    uiSB := mainGui.Add("StatusBar",, " 准备就绪。" modeText)

    mainGui.Show("w680 h670")
}

SaveSingleSetting(section, key, val) {
    SetConfigVal(section, key, val)
}

OnPageFormatChange(*) {
    global uiPageFormat, PageNumFormat
    PageNumFormat := uiPageFormat.Text
    SetConfigVal("PageNumber", "PageNumFormat", PageNumFormat)
}

OnFontSchemeChange(*) {
    global uiFontScheme, CurrentFontScheme
    CurrentFontScheme := uiFontScheme.Text
    SetConfigVal("FontScheme", "CurrentFontScheme", CurrentFontScheme)
    ApplyFontScheme(CurrentFontScheme)
}

; =====================================================================
; [自定义字体字号方案管理子窗口]
; =====================================================================
OpenFontSchemeManager(*) {
    global mainGui, CustomFontSchemes, uiFontScheme, CurrentFontScheme
    global WpsFontNameList, WpsFontSizeList, LineSpacingPresets

    mGui := Gui("+Owner" mainGui.Hwnd " +ToolWindow", "自定义字体字号方案管理")
    mGui.SetFont("s9", "Microsoft YaHei")

    mGui.Add("Text", "x20 y15 w70", "已有方案:")
    schemeItems := []
    for sName in CustomFontSchemes
        schemeItems.Push(sName)
    lbSchemes := mGui.Add("ListBox", "x20 y35 w150 h415", schemeItems)

    mGui.Add("GroupBox", "x185 y15 w520 h435", "方案参数配置 (各级文本对象)")
    mGui.Add("Text", "x200 y40 w65", "方案名称:")
    edName := mGui.Add("Edit", "x270 y37 w160")

    itemsDef := [
        ["主标题:",   "MTitle",   "方正小标宋简体", "二号", 0, "小标宋 二号 不加粗"],
        ["副标题:",   "SubTitle", "楷体",           "三号", 0, "楷体 三号 不加粗"],
        ["作者/发文:", "Author",   "楷体",           "三号", 0, "楷体 三号 不加粗"],
        ["正文内容:", "Body",     "仿宋",           "三号", 0, "仿宋 三号 不加粗"],
        ["一级标题:", "H1",       "黑体",           "三号", 0, "黑体 三号 不加粗"],
        ["二级标题:", "H2",       "楷体",           "三号", 0, "楷体 三号 不加粗"],
        ["三级标题:", "H3",       "仿宋",           "三号", 1, "仿宋 三号 (常加粗)"],
        ["四级标题:", "H4",       "仿宋",           "三号", 0, "仿宋 三号 不加粗"],
        ["表格内容:", "Table",    "仿宋",           "五号", 0, "仿宋 五号 不加粗"]
    ]

    ctrlMap := Map()
    curY := 70

    for it in itemsDef {
        mGui.Add("Text", "x200 y" curY + 3 " w65", it[1])
        cbF := mGui.Add("ComboBox", "x270 y" curY " w120", WpsFontNameList)
        cbF.Text := it[3]
        cbS := mGui.Add("ComboBox", "x395 y" curY " w75", WpsFontSizeList)
        cbS.Text := it[4]
        chkB := mGui.Add("Checkbox", "x478 y" curY + 3 " w50 Checked" it[5], "加粗")

        mGui.SetFont("s8 c777777")
        mGui.Add("Text", "x535 y" curY + 3 " w160", "（国标: " it[6] "）")
        mGui.SetFont("s9 c000000")

        ctrlMap[it[2] "Font"] := cbF
        ctrlMap[it[2] "Size"] := cbS
        ctrlMap[it[2] "Bold"] := chkB
        curY += 33
    }

    mGui.Add("Text", "x200 y" curY + 3 " w65", "正文行距:")
    cbSpacing := mGui.Add("ComboBox", "x270 y" curY " w160", LineSpacingPresets)
    cbSpacing.Text := "固定值 28 磅"
    mGui.SetFont("s8 c777777")
    mGui.Add("Text", "x440 y" curY + 3 " w250", "（未输入'倍'默认按磅，国标: 28 磅）")
    mGui.SetFont("s9 c000000")

    LoadSelectedSchemeToEditor() {
        sel := lbSchemes.Text
        if (sel != "" && CustomFontSchemes.Has(sel)) {
            data := CustomFontSchemes[sel]
            edName.Value := sel
            for k, ctrl in ctrlMap {
                if InStr(k, "Bold")
                    ctrl.Value := Integer(data.%k%)
                else
                    ctrl.Text := data.%k%
            }
            cbSpacing.Text := data.Spacing
        }
    }

    lbSchemes.OnEvent("Change", (*) => LoadSelectedSchemeToEditor())

    if (CurrentFontScheme != "党政机关公文格式" && CustomFontSchemes.Has(CurrentFontScheme)) {
        for idx, sName in schemeItems {
            if (sName == CurrentFontScheme) {
                lbSchemes.Choose(idx)
                LoadSelectedSchemeToEditor()
                break
            }
        }
    }

    btnSave := mGui.Add("Button", "x280 y460 w160 h32 Default", "保存/新建此方案")
    btnDelete := mGui.Add("Button", "x470 y460 w150 h32", "删除选中方案")

    OnSaveSchemeClick(*) {
        sName := Trim(edName.Value)
        if (sName == "" || sName == "党政机关公文格式") {
            MsgBox("请输入有效的方案名称，且不能与默认格式重名！", "提示", "Icon!")
            return
        }
        packObj := {Spacing: Trim(cbSpacing.Text)}
        for k, ctrl in ctrlMap
            packObj.%k% := InStr(k, "Bold") ? Integer(ctrl.Value) : Trim(ctrl.Text)

        SaveFontSchemeData(sName, packObj)
        CustomFontSchemes[sName] := packObj
        RefreshFontSchemeDropdown(sName)
        mGui.Destroy()
    }

    OnDeleteSchemeClick(*) {
        sel := lbSchemes.Text
        if (sel != "" && CustomFontSchemes.Has(sel)) {
            DeleteFontSchemeData(sel)
            CustomFontSchemes.Delete(sel)
            RefreshFontSchemeDropdown("党政机关公文格式")
            mGui.Destroy()
        }
    }

    btnSave.OnEvent("Click", OnSaveSchemeClick)
    btnDelete.OnEvent("Click", OnDeleteSchemeClick)
    mGui.Show("w720 h510")
}

RefreshFontSchemeDropdown(activeScheme := "党政机关公文格式") {
    global uiFontScheme, CustomFontSchemes, CurrentFontScheme
    newList := ["党政机关公文格式"]
    for sName in CustomFontSchemes
        newList.Push(sName)
    uiFontScheme.Delete()
    uiFontScheme.Add(newList)

    targetIdx := 1
    for idx, item in newList {
        if (item == activeScheme) {
            targetIdx := idx
            break
        }
    }
    uiFontScheme.Choose(targetIdx)
    CurrentFontScheme := activeScheme
    SetConfigVal("FontScheme", "CurrentFontScheme", CurrentFontScheme)
    ApplyFontScheme(CurrentFontScheme)
}

RestoreSelectedLVItem(*) {
    global uiLV
    row := uiLV.GetNext(0)
    if (row > 0) {
        path := uiLV.GetText(row, 3)
        if (RestoreLatestBackup(path))
            uiLV.Modify(row, "Col2", "已恢复备份")
    }
}

; 双击任务列表中的文件，用系统默认程序打开
OpenSelectedLVFile(*) {
    global uiLV, uiSB
    row := uiLV.GetNext(0)
    if (row <= 0)
        return
    filePath := uiLV.GetText(row, 3)
    if (filePath == "" || !FileExist(filePath)) {
        uiSB.SetText(" 文件不存在：" filePath)
        ShowFloatingToast("打开失败", "文件不存在或已被移动：`n" filePath, 3000, false)
        return
    }
    try {
        Run('"' filePath '"')
        uiSB.SetText(" 已打开：" filePath)
    } catch Error as err {
        uiSB.SetText(" 打开失败：" err.Message)
        ShowFloatingToast("打开失败", err.Message, 3000, false)
    }
}

CheckButtonHover(wParam, lParam, msg, hwnd) {
    global uiBtnKill, uiBakOpen
    if (IsSet(uiBakOpen) && IsObject(uiBakOpen) && hwnd == uiBakOpen.Hwnd) {
        ToolTip("打开备份文件夹")
        SetTimer(HideButtonToolTip, -1200)
        return
    }
    if (IsSet(uiBtnKill) && IsObject(uiBtnKill) && hwnd == uiBtnKill.Hwnd) {
        ToolTip("强制结束由本程序在后台调用的排版进程`n（安全操作，不影响您自行打开的文档）")
        SetTimer(() => ToolTip(), -300)
    }
}

HideButtonToolTip() {
    ToolTip()
}

GuiClose(*) {
    global BakOpenHIcon
    if (BakOpenHIcon) {
        DllCall("user32\DestroyIcon", "Ptr", BakOpenHIcon)
        BakOpenHIcon := 0
    }
    ExitApp()
}

ToggleBak(*) {
    global CreateBAK
    ; 主菜单项没有自动勾选行为，这里手动切换并同步配置文件
    CreateBAK := CreateBAK ? 0 : 1
    UpdateBakMenuCheck()
    SaveSingleSetting("Options", "CreateBAK", CreateBAK)
}

; 按当前 CreateBAK 值同步主菜单项的勾选状态
UpdateBakMenuCheck() {
    global uiSettingsMenu, CreateBAK, BAK_MENU_LABEL
    if (!IsSet(uiSettingsMenu) || !IsObject(uiSettingsMenu))
        return
    if (CreateBAK)
        uiSettingsMenu.Check(BAK_MENU_LABEL)
    else
        uiSettingsMenu.Uncheck(BAK_MENU_LABEL)
}

; 排版前确认备份开关：返回 true 表示可以继续排版
; 未开启备份时弹出“是/否”提醒，默认“是”（开启备份）；关闭或显示失败则视为取消
EnsureBackupConfirmed() {
    global CreateBAK
    if (CreateBAK)
        return true

    bakTip := MsgBox("检测到“排版前生成原始备份”当前处于关闭状态。`n`n"
        . "是否在开始排版前为每个文件生成原始备份？`n`n"
        . "· 选择“是”：开启备份并继续排版（推荐）`n"
        . "· 选择“否”：本次不备份，直接排版", "备份提醒", "YesNo Icon!")

    if (bakTip == "Yes") {
        CreateBAK := 1
        UpdateBakMenuCheck()
        SetConfigVal("Options", "CreateBAK", CreateBAK)
        return true
    }
    return (bakTip == "No")
}

ClearList(*) {
    global uiLV, uiSB
    uiLV.Delete()
    uiSB.SetText(" 列表已清空。")
}

KillAllWordProcesses(*) {
    global uiSB, ManagedWordPIDs
    killed := 0
    for pid in ManagedWordPIDs {
        if (ProcessExist(pid)) {
            try {
                ProcessClose(pid)
                killed++
                Sleep(100)
            } catch {
                continue
            }
        }
    }
    ManagedWordPIDs := Map()
    if (killed > 0)
        uiSB.SetText(" 清理完毕！共结束了 " killed " 个由本程序在后台调用的驻留进程。")
    else
        uiSB.SetText(" 当前系统未检测到由本程序调用的后台排版进程。")
}

EditSelectAll(ctrlObj, *) {
    SendMessage(0x00B1, 0, -1, ctrlObj.Hwnd)
}

AdjustMarginBySpinner(editCtrl, udCtrl) {
    global isProgramUpdating
    if (isProgramUpdating)
        return

    diff := udCtrl.Value - 50
    isProgramUpdating := true
    udCtrl.Value := 50
    isProgramUpdating := false

    currentVal := IsNumber(editCtrl.Value) ? Float(editCtrl.Value) : 0.0
    newVal := Max(0.0, Round(currentVal + diff * 1, 1))

    isProgramUpdating := true
    editCtrl.Value := (Round(newVal, 1) == Round(newVal)) ? String(Round(newVal)) : Format("{:.1f}", newVal)
    isProgramUpdating := false

    OnMarginEditChange()
}

OnMarginTypeChange(*) {
    global isProgramUpdating, CustomTop, CustomBottom, CustomLeft, CustomRight
    global uiMarginType, uiTop, uiBottom, uiLeft, uiRight

    isProgramUpdating := true
    sel := uiMarginType.Text
    if (sel == "党政机关公文格式") {
        uiTop.Value := "37", uiBottom.Value := "35", uiLeft.Value := "28", uiRight.Value := "26"
    } else if (sel == "WORD默认边距" || sel == "WORD默认页边距") {
        uiTop.Value := "25.4", uiBottom.Value := "25.4", uiLeft.Value := "31.8", uiRight.Value := "31.8"
    } else if (sel == "用户自定义") {
        uiTop.Value := (Round(CustomTop, 1) == Round(CustomTop)) ? String(Round(CustomTop)) : Format("{:.1f}", CustomTop)
        uiBottom.Value := (Round(CustomBottom, 1) == Round(CustomBottom)) ? String(Round(CustomBottom)) : Format("{:.1f}", CustomBottom)
        uiLeft.Value := (Round(CustomLeft, 1) == Round(CustomLeft)) ? String(Round(CustomLeft)) : Format("{:.1f}", CustomLeft)
        uiRight.Value := (Round(CustomRight, 1) == Round(CustomRight)) ? String(Round(CustomRight)) : Format("{:.1f}", CustomRight)
    }
    isProgramUpdating := false
    SaveMarginsRealtime()
}

OnMarginEditChange(*) {
    global isProgramUpdating, CustomTop, CustomBottom, CustomLeft, CustomRight
    global uiMarginType, uiTop, uiBottom, uiLeft, uiRight

    if (isProgramUpdating)
        return

    isProgramUpdating := true
    uiMarginType.Choose("用户自定义")
    isProgramUpdating := false

    CustomTop    := IsNumber(uiTop.Value) ? Float(uiTop.Value) : 0.0
    CustomBottom := IsNumber(uiBottom.Value) ? Float(uiBottom.Value) : 0.0
    CustomLeft   := IsNumber(uiLeft.Value) ? Float(uiLeft.Value) : 0.0
    CustomRight  := IsNumber(uiRight.Value) ? Float(uiRight.Value) : 0.0

    SaveMarginsRealtime()
}

SaveMarginsRealtime() {
    global uiMarginType, uiTop, uiBottom, uiLeft, uiRight, CustomTop, CustomBottom, CustomLeft, CustomRight
    SetConfigVal("Margins", "MarginType", uiMarginType.Value)
    SetConfigVal("Margins", "MarginTop", uiTop.Value)
    SetConfigVal("Margins", "MarginBottom", uiBottom.Value)
    SetConfigVal("Margins", "MarginLeft", uiLeft.Value)
    SetConfigVal("Margins", "MarginRight", uiRight.Value)
    SetConfigVal("Margins", "CustomTop", Format("{:.1f}", CustomTop))
    SetConfigVal("Margins", "CustomBottom", Format("{:.1f}", CustomBottom))
    SetConfigVal("Margins", "CustomLeft", Format("{:.1f}", CustomLeft))
    SetConfigVal("Margins", "CustomRight", Format("{:.1f}", CustomRight))
}

GuiDropFiles(guiObj, ctrlObj, fileArray, x, y) {
    global uiLV, uiSB
    existingPaths := Map()
    Loop uiLV.GetCount() {
        p := uiLV.GetText(A_Index, 3)
        if (p != "")
            existingPaths[Format("{:L}", p)] := true
    }

    ; 拖入的文件夹若包含子文件夹，先询问是否连同子文件夹一起扫描；默认按钮为“是”
    scanSubfolders := false
    for path in fileArray {
        if (InStr(FileExist(path), "D") && HasSubFolders(path)) {
            answer := MsgBox("检测到拖入的文件夹中包含子文件夹。`n`n"
                . "是否扫描子文件夹中的 Word 文件？`n`n"
                . "· 选择“是”：包含所有子文件夹中的文件（推荐）`n"
                . "· 选择“否”：只添加该文件夹第一层的文件", "扫描子文件夹", "YesNo Icon?")
            scanSubfolders := (answer == "Yes")
            break
        }
    }

    addedCount := 0, skippedCount := 0
    for path in fileArray {
        if InStr(FileExist(path), "D") {
            ; FR = 文件 + 递归子文件夹；F = 仅当前文件夹第一层
            fileOpts := scanSubfolders ? "FR" : "F"
            Loop Files, path "\*.*", fileOpts {
                ext := Format("{:L}", A_LoopFileExt)
                if ((ext == "doc" || ext == "docx") && SubStr(A_LoopFileName, 1, 2) != "~$") {
                    fullPathKey := Format("{:L}", A_LoopFilePath)
                    if (!existingPaths.Has(fullPathKey)) {
                        existingPaths[fullPathKey] := true
                        uiLV.Add("", A_LoopFileName, "等待排版", A_LoopFilePath)
                        addedCount++
                    } else {
                        skippedCount++
                    }
                }
            }
        } else {
            SplitPath(path, &fileName, , &fileExt)
            ext := Format("{:L}", fileExt)
            if ((ext == "doc" || ext == "docx") && SubStr(fileName, 1, 2) != "~$") {
                fullPathKey := Format("{:L}", path)
                if (!existingPaths.Has(fullPathKey)) {
                    existingPaths[fullPathKey] := true
                    uiLV.Add("", fileName, "等待排版", path)
                    addedCount++
                } else {
                    skippedCount++
                }
            }
        }
    }

    statusMsg := " 已添加 " addedCount " 个新任务。"
    if (skippedCount > 0)
        statusMsg .= " (已跳过 " skippedCount " 个重复文件)"
    statusMsg .= " 当前列表共 " uiLV.GetCount() " 个排版任务。"
    uiSB.SetText(statusMsg)
}

; 判断指定目录下是否存在子文件夹
HasSubFolders(dirPath) {
    Loop Files, dirPath "\*", "D"
        return true
    return false
}

SelectBakPath(*) {
    global uiBakPath
    selectedDir := BrowseFolderWithNewButton(uiBakPath.Value)
    if (selectedDir != "") {
        uiBakPath.Value := selectedDir
        SaveSingleSetting("Options", "BakDir", selectedDir)
    }
}

BrowseFolderWithNewButton(initialDir := "") {
    global DefaultBakDir
    ; 起始路径为空或已不存在时回退到“文档”，避免对话框定位失败
    startDir := Trim(initialDir)
    if (startDir == "")
        startDir := DefaultBakDir
    if (!DirExist(startDir))
        startDir := DirExist(A_MyDocuments) ? A_MyDocuments : ""

    ; 前缀 "*" 允许向上逐级导航（旧实现把路径当作根目录，导致无法切换到其它目录）；
    ; Options=3 同时提供“新建文件夹”按钮和可输入的编辑框
    try return DirSelect(startDir = "" ? "" : "*" startDir, 3, "请选择备份存放文件夹")
    catch
        return ""
}

; 在资源管理器中打开当前设置的备份目录
OpenBakFolder(*) {
    global uiBakPath, DefaultBakDir
    dir := Trim(uiBakPath.Value)
    if (dir == "")
        dir := DefaultBakDir
    if (!DirExist(dir)) {
        ; 目录尚未创建时自动创建，便于用户提前查看或管理备份位置
        try DirCreate(dir)
        catch {
            ShowFloatingToast("打开失败", "备份目录不存在且无法创建：`n" dir, 3000, false)
            return
        }
    }
    try Run('"' dir '"')
    catch {
        ShowFloatingToast("打开失败", "无法打开备份目录，请检查路径或权限！", 3000, false)
    }
}

; =====================================================================
; [系统标准图标：取 shell 的 SIID_FOLDEROPEN，并贴到方形图标按钮上]
; =====================================================================
GetStockIconHandle(iconID, flags) {
    ; SHSTOCKICONINFO：cbSize + 对齐 + hIcon + iSysImageIndex + iIcon + szPath[260]
    static infoSize := 4 + A_PtrSize + 4 + 4 + 520 + (A_PtrSize = 8 ? 4 : 0)
    info := Buffer(infoSize, 0)
    NumPut("UInt", infoSize, info, 0)
    if (DllCall("shell32\SHGetStockIconInfo", "Int", iconID, "UInt", flags, "Ptr", info, "Int") != 0)
        return 0
    return NumGet(info, A_PtrSize = 8 ? 8 : 4, "Ptr")
}

SetButtonStockIcon(btn, iconID := 4) {
    global BakOpenHIcon
    ; SHGSI_ICON(0x100) | SHGSI_SMALLICON(0x1)
    hIcon := GetStockIconHandle(iconID, 0x101)
    if (!hIcon)
        return false

    ; BS_ICON 按钮：BM_SETIMAGE(0x00F7) + IMAGE_ICON(1) 设置图标；
    ; BM_SETIMAGE 返回“之前的图标句柄”，首次设置返回 0 属正常，故不据此判断成败。
    DllCall("user32\SendMessageW", "Ptr", btn.Hwnd, "UInt", 0x00F7, "Ptr", 1, "Ptr", hIcon, "Ptr")

    ; 句柄需由本进程一直持有，直到按钮销毁，不能立即 DestroyIcon
    BakOpenHIcon := hIcon
    return true
}

ResetDefaults(*) {
    global uiMarginType, uiTop, uiBottom, uiLeft, uiRight, uiPageFormat, uiSkipSingle
    global uiFontScheme, CurrentFontScheme
    global uiPDF, uiFixDocNum, uiProtectTable, uiPreventOrphan
    global uiBakPath, uiBakBtn, uiBakOpen, uiSB, isProgramUpdating, DefaultBakDir, MaxHeadingLen, CreateBAK

    isProgramUpdating := true
    uiMarginType.Choose("党政机关公文格式")
    uiTop.Value := "37", uiBottom.Value := "35", uiLeft.Value := "28", uiRight.Value := "26"
    uiFontScheme.Choose("党政机关公文格式")
    CurrentFontScheme := "党政机关公文格式"
    SetConfigVal("FontScheme", "CurrentFontScheme", CurrentFontScheme)
    ApplyFontScheme(CurrentFontScheme)

    uiPageFormat.Choose(6)
    uiSkipSingle.Value := 1
    isProgramUpdating := false

    uiFixDocNum.Value := 0, uiProtectTable.Value := 0, uiPreventOrphan.Value := 0
    uiPDF.Value := 0
    ; 备份开关恢复为“开启”，并同步主菜单勾选状态
    CreateBAK := 1
    UpdateBakMenuCheck()
    uiBakPath.Value := DefaultBakDir
    uiBakPath.Enabled := 1, uiBakBtn.Enabled := 1, uiBakOpen.Enabled := 1
    MaxHeadingLen := 35

    SaveAllCurrentSettings()
    uiSB.SetText(" 参数已恢复为 GB/T 9704-2012 国家标准默认值（标题阈值：35字）。")
}

; =====================================================================
; [系统级辅助函数：屏幕右下角免焦点悬浮卡片]
; =====================================================================
DestroyCurrentToast() {
    global toastGui, toastTitleCtrl, toastMsgCtrl, toastIconCtrl
    SetTimer(DestroyCurrentToast, 0)
    if (toastGui != "" && IsObject(toastGui)) {
        try {
            if (WinExist(toastGui.Hwnd))
                toastGui.Destroy()
        }
        toastGui := "", toastTitleCtrl := "", toastMsgCtrl := "", toastIconCtrl := ""
    }
}

ShowFloatingToast(title, msg, duration := 2000, isSuccess := true, opacity := 220) {
    global toastGui, toastTitleCtrl, toastMsgCtrl, toastIconCtrl
    bgColor := isSuccess ? "1E392A" : "5C2020"
    iconText := isSuccess ? "✔" : "✖"

    ; 每次都重新创建，确保文字改动后控件高度能同步刷新
    DestroyCurrentToast()

    toastGui := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x08000000")
    toastGui.BackColor := bgColor
    toastGui.MarginX := 20, toastGui.MarginY := 14

    toastGui.SetFont("s12 bold cFFFFFF", "Microsoft YaHei")
    toastIconCtrl := toastGui.Add("Text", "x20 y14 w24", iconText)
    toastTitleCtrl := toastGui.Add("Text", "x+2 yp w294", title)

    toastGui.SetFont("s10 norm cE0E0E0", "Microsoft YaHei")
    toastMsgCtrl := toastGui.Add("Text", "x20 y+8 w320", msg)

    ; 依据消息文字的实际高度自适应窗口高度，避免长路径被截断
    toastW := 360
    toastMsgCtrl.GetPos(&msgX, &msgY, &msgW, &msgH)
    toastH := msgY + msgH + 14
    posX := A_ScreenWidth - toastW - 25
    posY := A_ScreenHeight - toastH - 55

    toastGui.Show("x" posX " y" posY " w" toastW " h" toastH " NoActivate")
    try WinSetTransparent(opacity, toastGui.Hwnd)

    if (duration > 0)
        SetTimer(DestroyCurrentToast, -duration)
}

WriteLogFile(success_list, fail_list, warn_list := []) {
    ; A_ScriptDir 编译后为 EXE 所在目录，未编译时为脚本所在目录
    ; 日志固定为一个文件，每次运行覆盖写入，文件名不带日期
    log_path := A_ScriptDir "\公文排版执行日志.txt"

    log_content := "====================================================`r`n"
                 . "         党政机关公文自动排版执行日志`r`n"
                 . "  生成时间: " FormatTime(, "yyyy-MM-dd HH:mm:ss") "`r`n"
                 . "====================================================`r`n`r`n"
                 . "【任务概况】`r`n"
                 . "  排版成功数量: " success_list.Length " 个`r`n"
                 . "  排版失败数量: " fail_list.Length " 个`r`n"
                 . "  成功但有提示: " warn_list.Length " 个`r`n`r`n"
                 . "----------------------------------------------------`r`n"
                 . "【失败列表】 (共 " fail_list.Length " 个，可能文件被占用或损坏)`r`n"
                 . "----------------------------------------------------`r`n"
    if (fail_list.Length == 0) {
        log_content .= "  (无失败记录)`r`n"
    } else {
        for idx, item in fail_list
            log_content .= idx ". 文件路径: " item.path "`r`n   失败原因: " item.err "`r`n`r`n"
    }

    log_content .= "----------------------------------------------------`r`n"
                 . "【排版成功但有提示】 (共 " warn_list.Length " 个，请人工复核)`r`n"
                 . "----------------------------------------------------`r`n"
    if (warn_list.Length == 0) {
        log_content .= "  (无提示记录)`r`n"
    } else {
        for idx, item in warn_list
            log_content .= idx ". 文件路径: " item.path "`r`n   提示信息: " item.err "`r`n`r`n"
    }

    try {
        file_obj := FileOpen(log_path, "w", "UTF-8")
        file_obj.Write(log_content)
        file_obj.Close()
        return log_path
    } catch {
        return ""
    }
}

MmToPt(mm) => Float(mm) * 2.83465

; =====================================================================
; [加载配置与注册表右键菜单管理]
; =====================================================================
LoadSettings() {
    global MarginType := GetConfigVal("Margins", "MarginType", 1)
    global MarginTop := GetConfigVal("Margins", "MarginTop", 37)
    global MarginBottom := GetConfigVal("Margins", "MarginBottom", 35)
    global MarginLeft := GetConfigVal("Margins", "MarginLeft", 28)
    global MarginRight := GetConfigVal("Margins", "MarginRight", 26)

    global CustomTop := GetConfigVal("Margins", "CustomTop", 0.0)
    global CustomBottom := GetConfigVal("Margins", "CustomBottom", 0.0)
    global CustomLeft := GetConfigVal("Margins", "CustomLeft", 0.0)
    global CustomRight := GetConfigVal("Margins", "CustomRight", 0.0)

    global PageNumFormat := GetConfigVal("PageNumber", "PageNumFormat", "格式数字 — 1 — 居中")
    global SkipSinglePageNum := GetConfigVal("PageNumber", "SkipSinglePageNum", 1)
    global CurrentFontScheme := GetConfigVal("FontScheme", "CurrentFontScheme", "党政机关公文格式")

    global ExportPDF := GetConfigVal("Options", "ExportPDF", 0)
    global CreateBAK := GetConfigVal("Options", "CreateBAK", 1)
    global OptFixDocNum := GetConfigVal("Options", "OptFixDocNum", 0)
    global OptProtectTable := GetConfigVal("Options", "OptProtectTable", 0)
    global OptPreventOrphan := GetConfigVal("Options", "OptPreventOrphan", 0)
    global MaxHeadingLen := Integer(GetConfigVal("Options", "MaxHeadingLen", 35))
    global ContextActive := Integer(GetConfigVal("Options", "ContextActive", 0))

    if (MaxHeadingLen == 50) {
        MaxHeadingLen := 35
        SetConfigVal("Options", "MaxHeadingLen", 35)
    }

    global BakDir := GetConfigVal("Options", "BakDir", DefaultBakDir)
    if (BakDir == "")
        BakDir := DefaultBakDir

    LoadFontSchemes()
    ApplyFontScheme(CurrentFontScheme)
}

CreateTraySettingsMenu() {
    global SettingsMenu
    try A_TrayMenu.Delete()

    A_TrayMenu.Add("打开主窗口", (*) => ShowGUI())
    A_TrayMenu.Add("设置(&S)", SettingsMenu := Menu())
    SettingsMenu.Add("添加 Word 文件右键运行命令", (*) => RegisterDocumentContextMenu(true))
    SettingsMenu.Add("删除 Word 文件右键运行命令", (*) => RemoveDocumentContextMenu())
    A_TrayMenu.Add()
    A_TrayMenu.Add("退出(&X)", (*) => ExitApp())
}

GetContextMenuBaseClasses() {
    return [
        "HKCU\Software\Classes\.doc\shell",
        "HKCU\Software\Classes\.docx\shell",
        "HKCU\Software\Classes\SystemFileAssociations\.doc\Shell",
        "HKCU\Software\Classes\SystemFileAssociations\.docx\Shell",
        "HKCU\Software\Classes\Word.Document.12\shell",
        "HKCU\Software\Classes\Word.DocumentMacroEnabled.12\shell",
        "HKCU\Software\Classes\Word.Document.8\shell",
        "HKCU\Software\Classes\Word.DocumentMacroEnabled.8\shell"
    ]
}

RegisterDocumentContextMenu(showToast := true) {
    if (A_IsCompiled) {
        cmdPrefix := '"' . A_ScriptFullPath . '"', iconPath := A_ScriptFullPath
    } else if FileExist(A_ScriptDir "\公文自动排版.exe") {
        defaultExe := A_ScriptDir "\公文自动排版.exe"
        cmdPrefix := '"' . defaultExe . '"', iconPath := defaultExe
    } else {
        cmdPrefix := '"' . A_AhkPath . '" "' . A_ScriptFullPath . '"', iconPath := A_AhkPath
    }

    menuRunTitle     := "公文自动排版(&Z)"
    commandRun       := cmdPrefix . ' "%1"'

    hasError := false
    for _, basePath in GetContextMenuBaseClasses() {
        try {
            runKey := basePath "\DocFormatToolRun"
            RegCreateKey(runKey "\command")
            RegWrite(menuRunTitle, "REG_SZ", runKey)
            RegWrite(menuRunTitle, "REG_SZ", runKey, "MUIVerb")
            RegWrite(iconPath, "REG_SZ", runKey, "Icon")
            RegWrite(commandRun, "REG_SZ", runKey "\command")

            ; 清理旧版可能遗留的恢复菜单项
            try RegDeleteKey(basePath "\DocFormatToolRestore")
        } catch {
            hasError := true
        }
    }

    global ContextActive := 1
    SetConfigVal("Options", "ContextActive", 1)

    if (showToast) {
        if (!hasError) {
            ShowFloatingToast("右键命令添加成功", "已注册【" menuRunTitle "】命令", 2500, true)
        } else {
            ShowFloatingToast("注册提示", "部分项目写入受阻，请检查权限", 3500, false)
        }
    }
    return !hasError
}

RemoveDocumentContextMenu() {
    global ContextActive := 0
    SetConfigVal("Options", "ContextActive", 0)

    for _, basePath in GetContextMenuBaseClasses() {
        try RegDeleteKey(basePath "\DocFormatToolRun")
        try RegDeleteKey(basePath "\DocFormatToolRestore")
    }
    ShowFloatingToast("操作完成", "已删除 Word 文件右键菜单命令", 3000, true)
    return true
}

CheckAndRepairContextMenuPath() {
    regKeyPath := "HKCU\Software\Classes\.docx\shell\DocFormatToolRun\command"
    try {
        regCmd := RegRead(regKeyPath, "")
        regExePath := Trim(RegExReplace(regCmd, '"\s*"%1"$', ""), '" ')
        if (regExePath != "" && !FileExist(regExePath)) {
            RegisterDocumentContextMenu(false)
        }
    } catch {
        RegisterDocumentContextMenu(false)
    }
}

SaveAllCurrentSettings() {
    global MarginType := uiMarginType.Value
    global MarginTop := uiTop.Value
    global MarginBottom := uiBottom.Value
    global MarginLeft := uiLeft.Value
    global MarginRight := uiRight.Value
    global PageNumFormat := uiPageFormat.Text
    global SkipSinglePageNum := uiSkipSingle.Value
    global CurrentFontScheme := uiFontScheme.Text
    global ExportPDF := uiPDF.Value
    global CreateBAK
    global OptFixDocNum := uiFixDocNum.Value
    global OptProtectTable := uiProtectTable.Value
    global OptPreventOrphan := uiPreventOrphan.Value
    global BakDir := uiBakPath.Value
    global MaxHeadingLen

    if (MarginType == 2) {
        MarginTop    := Format("{:.1f}", Float(MarginTop))
        MarginBottom := Format("{:.1f}", Float(MarginBottom))
        MarginLeft   := Format("{:.1f}", Float(MarginLeft))
        MarginRight  := Format("{:.1f}", Float(MarginRight))
    }

    SetConfigVal("Margins", "MarginType", MarginType)
    SetConfigVal("Margins", "MarginTop", MarginTop)
    SetConfigVal("Margins", "MarginBottom", MarginBottom)
    SetConfigVal("Margins", "MarginLeft", MarginLeft)
    SetConfigVal("Margins", "MarginRight", MarginRight)
    SetConfigVal("PageNumber", "PageNumFormat", PageNumFormat)
    SetConfigVal("PageNumber", "SkipSinglePageNum", SkipSinglePageNum)
    SetConfigVal("FontScheme", "CurrentFontScheme", CurrentFontScheme)
    SetConfigVal("Options", "ExportPDF", ExportPDF)
    SetConfigVal("Options", "CreateBAK", CreateBAK)
    SetConfigVal("Options", "OptFixDocNum", OptFixDocNum)
    SetConfigVal("Options", "OptProtectTable", OptProtectTable)
    SetConfigVal("Options", "OptPreventOrphan", OptPreventOrphan)
    SetConfigVal("Options", "BakDir", BakDir)
    SetConfigVal("Options", "MaxHeadingLen", MaxHeadingLen)

    ApplyFontScheme(CurrentFontScheme)
}

; =====================================================================
; [主排版流水线控制 (模式 B：GUI 列表批量排版)]
; =====================================================================
SetGUIControlsLock(isEnabled) {
    global uiBtnRun, uiBtnReset, uiBtnClear, uiBtnKill, uiBtnExport, uiTop, uiBottom, uiLeft, uiRight, uiPageFormat, uiSkipSingle
    global uiFontScheme, uiBtnManageScheme, uiPDF, uiFixDocNum, uiProtectTable, uiPreventOrphan, uiBakPath, uiBakBtn, uiBakOpen, uiMarginType
    global uiTopUD, uiBottomUD, uiLeftUD, uiRightUD, uiSettingsMenu, BAK_MENU_LABEL

    for ctrl in [uiBtnRun, uiBtnReset, uiBtnClear, uiBtnExport, uiMarginType, uiFontScheme, uiBtnManageScheme,
                 uiPageFormat, uiSkipSingle, uiTop, uiBottom, uiLeft, uiRight, uiTopUD, uiBottomUD, uiLeftUD, uiRightUD,
                 uiPDF, uiFixDocNum, uiProtectTable, uiPreventOrphan] {
        ctrl.Enabled := isEnabled
    }
    uiBtnKill.Enabled := true
    ; 备份目录控件和备份开关仅在批量排版期间锁定，平时保持可用
    uiBakPath.Enabled := isEnabled
    uiBakBtn.Enabled  := isEnabled
    uiBakOpen.Enabled := isEnabled
    if (IsSet(uiSettingsMenu) && IsObject(uiSettingsMenu)) {
        if (isEnabled)
            uiSettingsMenu.Enable(BAK_MENU_LABEL)
        else
            uiSettingsMenu.Disable(BAK_MENU_LABEL)
    }
}

RunGUIProcessing(*) {
    global uiLV, uiSB, CreateBAK
    count := uiLV.GetCount()
    if (count == 0) {
        uiSB.SetText(" 任务列表为空！请先将 Word 文件或文件夹拖入列表中。")
        return
    }

    ; 排版会直接修改原文件，备份开关关闭时先给出明确提醒；默认按钮为“是”（备份）
    if (!EnsureBackupConfirmed()) {
        uiSB.SetText(" 已取消本次排版：备份开关未开启。")
        return
    }

    SaveAllCurrentSettings()
    uiSB.SetText(" 参数已保存。正在初始化运行环境...")
    SetGUIControlsLock(false)
    
    app := GetWordApplication(&appName)
    if (!app) {
        uiSB.SetText(" 运行失败：未检测到可用的 WPS 或 Microsoft Word！")
        SetGUIControlsLock(true)
        return
    }

    success_list := [], fail_list := [], warn_list := []

    Loop count {
        row := A_Index
        filePath := uiLV.GetText(row, 3)
        fileName := uiLV.GetText(row, 1)

        uiLV.Modify(row, "Col2", "▶ 正在排版...")
        uiSB.SetText(" 正在排版 (" row "/" count "): " fileName)
        Sleep(-1)

        if (!IsAppAlive(app))
            app := GetWordApplication(&appName)

        err_msg := "", warn_msg := ""
        if (FormatWordDocument(filePath, app, &err_msg, &warn_msg)) {
            status_text := "✔ 成功"
            if (warn_msg != "") {
                if InStr(warn_msg, "PDF")
                    status_text .= "（PDF 失败）"
                else
                    status_text .= "（有提示）"
                warn_list.Push({path: filePath, err: warn_msg})
            }
            uiLV.Modify(row, "Col2", status_text)
            success_list.Push(filePath)
        } else {
            uiLV.Modify(row, "Col2", "❌ 失败: " err_msg)
            fail_list.Push({path: filePath, err: err_msg})
        }
        Sleep(-1)
    }

    if (IsSet(app) && app != "") {
        try {
            if (app.Documents.Count == 0)
                app.Quit(0)
        }
    }

    if (fail_list.Length > 0 || warn_list.Length > 0) {
        logPath := WriteLogFile(success_list, fail_list, warn_list)
        logHint := (logPath != "") ? "详情见程序所在目录日志" : "日志写入失败，请检查程序所在目录写入权限"
        pdf_warn_count := 0
        for w in warn_list {
            if InStr(w.err, "PDF")
                pdf_warn_count += 1
        }
        msg := "排版结束！成功: " success_list.Length " 个，失败: " fail_list.Length " 个"
        if (warn_list.Length > 0)
            msg .= "，有提示: " warn_list.Length " 个"
        if (pdf_warn_count > 0)
            msg .= "（PDF 失败 " pdf_warn_count " 个）"
        msg .= "。"
        uiSB.SetText(" " msg " " logHint "。")
        toast_title := (fail_list.Length > 0) ? "文件排版完毕 (有失败)" : "文件排版完毕 (有提示)"
        ShowFloatingToast(toast_title, msg "`n" logHint, 3500, false)
    } else {
        msg := "全部处理成功！共完成 " success_list.Length " 个文件。"
        uiSB.SetText(" " msg)
        ShowFloatingToast("文件排版完毕", msg, 2000, true)
    }

    SetGUIControlsLock(true)
}

; =====================================================================
; [主排版流水线控制 (模式 A：静默多选 / 外部右键菜单排版)]
; =====================================================================
RunSilentProcessing(target_files) {
    unique_map := Map(), clean_targets := []
    for f in target_files {
        if (!unique_map.Has(f) && f != "") {
            unique_map[f] := true
            clean_targets.Push(f)
        }
    }

    processing_files := []
    for path in clean_targets {
        if InStr(FileExist(path), "D") {
            Loop Files, path "\*.*", "F" {
                ext := Format("{:L}", A_LoopFileExt)
                if ((ext == "doc" || ext == "docx") && SubStr(A_LoopFileName, 1, 2) != "~$")
                    processing_files.Push(A_LoopFilePath)
            }
        } else {
            SplitPath(path, &fileName, , &fileExt)
            ext := Format("{:L}", fileExt)
            if ((ext == "doc" || ext == "docx") && SubStr(fileName, 1, 2) != "~$")
                processing_files.Push(path)
        }
    }

    total_count := processing_files.Length
    if (total_count == 0)
        ExitApp()

    ; 静默排版同样受备份开关保护：关闭时先询问（默认“是”）
    if (!EnsureBackupConfirmed())
        ExitApp()

    app := GetWordApplication(&appName)
    if (!app) {
        ShowFloatingToast("排版引擎错误", "未检测到可用的 WPS 或 Word！", 3500, false)
        Sleep(3500)
        ExitApp()
    }

    success_list := [], fail_list := [], warn_list := []

    for idx, filePath in processing_files {
        SplitPath(filePath, &fileName)
        ShowFloatingToast("正在排版 (" idx "/" total_count ")", fileName, 0, true)

        if (!IsAppAlive(app))
            app := GetWordApplication(&appName)

        err_msg := "", warn_msg := ""
        if (FormatWordDocument(filePath, app, &err_msg, &warn_msg)) {
            success_list.Push(filePath)
            if (warn_msg != "")
                warn_list.Push({path: filePath, err: warn_msg})
        } else {
            fail_list.Push({path: filePath, err: err_msg})
        }
    }

    if (IsSet(app) && app != "") {
        try {
            if (app.Documents.Count == 0)
                app.Quit(0)
        }
    }

    if (fail_list.Length > 0 || warn_list.Length > 0) {
        logPath := WriteLogFile(success_list, fail_list, warn_list)
        logHint := (logPath != "") ? "详情日志已保存在程序所在目录" : "日志写入失败，请检查程序所在目录写入权限"
        pdf_warn_count := 0
        for w in warn_list {
            if InStr(w.err, "PDF")
                pdf_warn_count += 1
        }
        msg := "成功: " success_list.Length " 个，失败: " fail_list.Length " 个"
        if (warn_list.Length > 0)
            msg .= "，有提示: " warn_list.Length " 个"
        if (pdf_warn_count > 0)
            msg .= "（PDF 失败 " pdf_warn_count " 个）"
        msg .= "`n" logHint
        toast_title := (fail_list.Length > 0) ? "文件排版完毕 (有失败)" : "文件排版完毕 (有提示)"
        ShowFloatingToast(toast_title, msg, 3500, false)
        Sleep(3500)
    } else {
        msg := "已顺利完成 " success_list.Length " 个文件的自动排版。"
        ShowFloatingToast("文件自动排版成功", msg, 2000, true)
        Sleep(2000)
    }
    ExitApp()
}

; =====================================================================
; [底层排版辅助函数区]
; =====================================================================

GetVisualLength(text) {
    len := 0.0
    Loop Parse text
        len += (Ord(A_LoopField) <= 127) ? 0.5 : 1.0
    return len
}

IsAppAlive(app) {
    try {
        test := app.Version
        return true
    } catch {
        return false
    }
}

GetWordApplication(&appName) {
    global ManagedWordPIDs
    candidates := [
        ["Kwps.Application", "WPS 文字 (Kwps)"],
        ["WPS.Application", "WPS 文字 (WPS)"],
        ["Word.Application", "Microsoft Word"]
    ]
    for item in candidates {
        try {
            app := ComObject(item[1])
            app.Visible := false
            try app.DisplayAlerts := 0
            appName := item[2]
            
            try {
                prev_Detect := A_DetectHiddenWindows
                DetectHiddenWindows(true)
                original_caption := app.Caption
                unique_caption := "DocFormatTool_Instance_" A_TickCount
                app.Caption := unique_caption
                if (hw := WinExist(unique_caption)) {
                    if (pid := WinGetPID("ahk_id " hw))
                        ManagedWordPIDs[pid] := true
                }
                app.Caption := original_caption
                DetectHiddenWindows(prev_Detect)
            }
            return app
        }
    }
    appName := ""
    return ""
}

SetRangeFont(rng, fName, fSize, fBold) {
    f := rng.Font
    f.NameFarEast := fName
    f.NameAscii   := fName
    f.Size        := fSize
    f.Bold        := fBold
}

SetParagraphText(paragraph, newText) {
    r := paragraph.Range.Duplicate
    if (r.End > r.Start)
        r.End := r.End - 1
    r.Text := newText
}

TrimAllSpaces(str) {
    return RegExReplace(str, "^[\s\x{00A0}\x{2000}-\x{200D}\x{FEFF}\x{3000}]+|[\s\x{00A0}\x{2000}-\x{200D}\x{FEFF}\x{3000}]+$", "")
}

LTrimAllSpaces(str) {
    return RegExReplace(str, "^[\s\x{00A0}\x{2000}-\x{200D}\x{FEFF}\x{3000}]+", "")
}

IsEndWithPunct(text) {
    clean_text := TrimAllSpaces(text)
    return (clean_text != "") && (InStr("。！？；：，、!?;:,", SubStr(clean_text, -1)) > 0)
}

IsEndWithColon(text) {
    clean_text := TrimAllSpaces(text)
    return (clean_text != "") && (InStr("：:", SubStr(clean_text, -1)) > 0)
}

GetCleanText(paragraph) {
    raw := RTrim(paragraph.Range.Text, "`r`n")
    cleaned := Trim(raw, " `t`v" . Chr(12288) . Chr(160))
    cleaned := RegExReplace(cleaned, "^[\x{2000}-\x{200D}\x{FEFF}]+|[\x{2000}-\x{200D}\x{FEFF}]+$", "")
    if (raw != cleaned && cleaned != "") {
        try {
            if (paragraph.Range.InlineShapes.Count == 0 && paragraph.Range.Fields.Count == 0) {
                SetParagraphText(paragraph, cleaned)
            } else {
                leadLen := StrLen(raw) - StrLen(LTrim(raw, " `t`v" . Chr(12288) . Chr(160)))
                if (leadLen > 0) {
                    rLead := paragraph.Range.Duplicate
                    rLead.SetRange(rLead.Start, rLead.Start + leadLen)
                    rLead.Delete()
                }
            }
        }
    }
    return cleaned
}

CleanAllBlankParagraphs(doc) {
    count := doc.Paragraphs.Count
    Loop count {
        if (Mod(A_Index, 20) == 0)
            Sleep(-1)
        if (doc.Paragraphs.Count <= 1)
            break
        i := count - A_Index + 1
        p := doc.Paragraphs(i)
        try {
            if (p.Range.Tables.Count > 0 || p.Range.InlineShapes.Count > 0)
                continue
        }
        txt := TrimAllSpaces(RTrim(p.Range.Text, "`r`n"))
        if (txt == "") {
            try p.Range.Delete()
        }
    }
}

RemoveTrailingEmptyParagraphs(doc) {
    while (doc.Paragraphs.Count > 1) {
        last_p := doc.Paragraphs(doc.Paragraphs.Count)
        try {
            if (last_p.Range.Tables.Count > 0 || last_p.Range.InlineShapes.Count > 0)
                break
        }
        txt := TrimAllSpaces(RTrim(last_p.Range.Text, "`r`n"))
        if (txt == "") {
            try last_p.Range.Delete()
            catch
                break
        } else {
            break
        }
    }
}

; ---------------------------------------------------------------------
; [页码安全写入辅助函数]
; 说明：只替换本程序书签标记的页码块；页脚中的其他文字、表格、图片
;       和用户自行插入的页码均保留，无法安全处理时记录提示。
; ---------------------------------------------------------------------
GetFooterVisibleText(obj) {
    txt := obj.Range.Text
    txt := StrReplace(txt, Chr(13), "")
    txt := StrReplace(txt, Chr(10), "")
    txt := StrReplace(txt, Chr(12), "")
    return TrimAllSpaces(txt)
}

GetFooterPageFieldInfo(footer) {
    pageCount := 0, numPagesCount := 0, otherCount := 0
    fields := footer.Range.Fields
    Loop fields.Count {
        f := fields.Item(A_Index)
        if (f.Type == 33)
            pageCount += 1
        else if (f.Type == 26)
            numPagesCount += 1
        else
            otherCount += 1
    }
    return {page: pageCount, numPages: numPagesCount, other: otherCount}
}

IsFooterOnlyPageNumber(footer) {
    info := GetFooterPageFieldInfo(footer)
    if (info.other > 0 || info.page == 0 || info.page > 1 || info.numPages > 1)
        return false

    ; 页脚含表格、图片、内容控件、超链接、书签、批注或修订时，绝不整段清空
    if (footer.Range.Tables.Count > 0)
        return false
    if (footer.Range.InlineShapes.Count > 0)
        return false
    if (footer.Range.Bookmarks.Count > 0)
        return false
    if (footer.Range.Hyperlinks.Count > 0)
        return false
    try {
        if (footer.Range.ShapeRange.Count > 0)
            return false
    }
    try {
        if (footer.Range.ContentControls.Count > 0)
            return false
    }
    try {
        if (footer.Range.Comments.Count > 0)
            return false
    }
    try {
        if (footer.Range.Revisions.Count > 0)
            return false
    }

    txt := GetFooterVisibleText(footer)
    ; 兼容本程序生成的“— 1 —”“1 / 5”“1”以及常见的“第 1 页 共 5 页”；
    ; 字段结果暂时为空时也允许通过，避免未更新域时误判。
    if RegExMatch(txt, "^\s*[—–\-－]?\s*\d*\s*[—–\-－]?\s*$")
        return true
    if RegExMatch(txt, "^\s*\d*\s*[/／]\s*\d*\s*$")
        return true
    if RegExMatch(txt, "^\s*第\s*\d*\s*页\s*(共\s*\d*\s*页)?\s*$")
        return true
    return false
}

FooterHasPageField(footer) {
    fields := footer.Range.Fields
    Loop fields.Count {
        f := fields.Item(A_Index)
        if (f.Type == 33 || f.Type == 26)
            return true
    }
    return false
}

UpdateFooterPageFields(footer) {
    fields := footer.Range.Fields
    Loop fields.Count {
        f := fields.Item(A_Index)
        if (f.Type == 33 || f.Type == 26)
            try f.Update()
    }
}

FooterTypeLabel(type) {
    if (type == 2)
        return "首页页脚"
    if (type == 3)
        return "偶数页页脚"
    return "页脚"
}

RemoveFooterPageBlock(footer) {
    ; 清理本程序写入的页码书签块。除精确书签名外，也清理因取消
    ; “链接到前一节”而被复制的书签，避免重复页码。
    bms := footer.Range.Bookmarks
    removed := false
    guard := 0
    i := 1
    while (i <= bms.Count) {
        guard += 1
        if (guard > 100)
            throw Error("页脚页码书签清理失败")
        name := bms.Item(i).Name
        if (SubStr(name, 1, 10) == "DocFmt_PN_") {
            bm := bms.Item(i)
            try bm.Range.Delete()
            try bm.Delete()
            removed := true
            i := 1
            continue
        }
        i += 1
    }
    return removed
}

WriteFooterPageBlock(footer, paraIndex, formatStr, bmName) {
    fontPageNum := "宋体", sizePageNum := 14
    alignVal := InStr(formatStr, "居左") ? 0 : (InStr(formatStr, "居右") ? 2 : 1)
    wdFieldPage := 33, wdFieldNumPages := 26

    target := footer.Range.Paragraphs.Item(paraIndex).Range
    r := target.Duplicate
    r.Collapse(1)

    if (InStr(formatStr, "格式数字")) {
        r.Text := " —"
        r.Collapse(1)
        footer.Range.Fields.Add(r, wdFieldPage, , false)
        rStart := target.Duplicate
        rStart.Collapse(1)
        rStart.Text := "— "
    } else if (InStr(formatStr, "加总页码")) {
        r.Text := " / "
        r.Collapse(0)
        footer.Range.Fields.Add(r, wdFieldNumPages, , false)
        rStart := target.Duplicate
        rStart.Collapse(1)
        footer.Range.Fields.Add(rStart, wdFieldPage, , false)
    } else {
        footer.Range.Fields.Add(r, wdFieldPage, , false)
    }

    ; 重新获取段落范围，确保包含刚插入的域和文字
    target := footer.Range.Paragraphs.Item(paraIndex).Range
    target.ParagraphFormat.Alignment := alignVal
    target.ParagraphFormat.CharacterUnitFirstLineIndent := 0
    target.ParagraphFormat.FirstLineIndent := 0
    target.ParagraphFormat.CharacterUnitLeftIndent := 0
    target.ParagraphFormat.LeftIndent := 0
    target.ParagraphFormat.CharacterUnitRightIndent := 0
    target.ParagraphFormat.RightIndent := 0
    target.ParagraphFormat.LineSpacingRule := 0 ; 0 = 单倍行距
    target.ParagraphFormat.SpaceBefore := 0
    target.ParagraphFormat.SpaceAfter := 0
    SetRangeFont(target, fontPageNum, sizePageNum, false)

    if (footer.Range.Bookmarks.Exists(bmName))
        footer.Range.Bookmarks.Item(bmName).Delete()
    footer.Range.Bookmarks.Add(bmName, target)
}

SetPageNumbers(doc, formatStr := "格式数字 — 1 — 居中", skipSinglePage := 1) {
    warnings := []
    totalPages := doc.ComputeStatistics(2)
    disable := ((totalPages <= 1 && skipSinglePage) || InStr(formatStr, "无"))

    for section in doc.Sections {
        pageSetup := section.PageSetup
        footerTypes := [1]
        if (pageSetup.DifferentFirstPageHeaderFooter)
            footerTypes.Push(2)
        if (pageSetup.OddAndEvenPagesHeaderFooter)
            footerTypes.Push(3)

        for footerType in footerTypes {
            bmName := "DocFmt_PN_S" section.Index "_T" footerType "_" A_TickCount
            try {
                footer := section.Footers(footerType)
                footer.LinkToPrevious := false

                ; 先移除上一次写入的页码块，保留页脚其他内容
                removed := RemoveFooterPageBlock(footer)
                UpdateFooterPageFields(footer)

                if (disable) {
                    if (removed)
                        continue
                    ; 仅在页脚确实“只含页码”时清除，避免误删其他页脚内容
                    if (IsFooterOnlyPageNumber(footer))
                        footer.Range.Text := ""
                    else if (FooterHasPageField(footer))
                        warnings.Push("第 " section.Index " 节" FooterTypeLabel(footerType) "含其他内容，未自动移除页码，请人工检查")
                    continue
                }

                ; 页脚只含旧版页码：可安全整体替换
                if (IsFooterOnlyPageNumber(footer)) {
                    footer.Range.Text := ""
                    WriteFooterPageBlock(footer, 1, formatStr, bmName)
                    continue
                }

                ; 页脚已有页码且还有其他内容：保留原内容，只统一页码域字体
                if (FooterHasPageField(footer)) {
                    fields := footer.Range.Fields
                    Loop fields.Count {
                        f := fields.Item(A_Index)
                        if (f.Type == 33 || f.Type == 26)
                            SetRangeFont(f.Result, "宋体", 14, false)
                    }
                    warnings.Push("第 " section.Index " 节" FooterTypeLabel(footerType) "含其他内容，已保留原有页码和页脚内容")
                    continue
                }

                visibleText := GetFooterVisibleText(footer)
                if (visibleText == "" && footer.Range.Tables.Count == 0) {
                    ; 页脚为空（或仅含图片等无文字内容），在第一个段落写入
                    WriteFooterPageBlock(footer, 1, formatStr, bmName)
                } else {
                    ; 页脚有其他内容但没有页码：在末尾追加独立段落，保留原内容
                    lastIdx := footer.Range.Paragraphs.Count
                    lastPara := footer.Range.Paragraphs.Item(lastIdx)
                    lastText := GetFooterVisibleText(lastPara)
                    lastParaHasObjects := (lastPara.Range.Bookmarks.Count > 0 || lastPara.Range.InlineShapes.Count > 0 || footer.Range.Tables.Count > 0)
                    try {
                        if (lastPara.Range.ShapeRange.Count > 0 || lastPara.Range.ContentControls.Count > 0)
                            lastParaHasObjects := true
                    }
                    if (lastText != "" || lastParaHasObjects) {
                        ; 在最后一个段落标记之前插入新段落，确保页码位于原有内容之后。
                        ; 不能直接对 footer.Range 调用 InsertParagraphAfter，否则单段页脚
                        ; 不会新增段落，页码会写到原段落开头。
                        endRange := footer.Range.Duplicate
                        endRange.SetRange(endRange.End - 1, endRange.End - 1)
                        endRange.InsertParagraphAfter()
                        lastIdx := footer.Range.Paragraphs.Count
                    }
                    WriteFooterPageBlock(footer, lastIdx, formatStr, bmName)
                }
            } catch Error as err {
                warnings.Push("第 " section.Index " 节" FooterTypeLabel(footerType) "处理失败：" err.Message)
            }
        }
    }

    if (warnings.Length == 0)
        return ""
    warnText := ""
    for w in warnings {
        if (warnText != "")
            warnText .= "；"
        warnText .= w
    }
    return warnText
}

ExportToPdf(doc, doc_path) {
    SplitPath(doc_path, , &dir, , &name_no_ext)
    pdf_path := dir "\" name_no_ext ".pdf"
    try {
        if (InStr(FileExist(pdf_path), "D"))
            return {ok: false, path: pdf_path, err: "同名 PDF 路径被文件夹占用"}
        if FileExist(pdf_path) {
            ; 先检测同名 PDF 是否被占用，避免导出时静默失败
            try {
                f := FileOpen(pdf_path, "a")
                f.Close()
            } catch {
                return {ok: false, path: pdf_path, err: "同名 PDF 已存在且被占用，请关闭后重试"}
            }
        }
        doc.ExportAsFixedFormat(pdf_path, 17, false, 0, 0)
        if (!FileExist(pdf_path))
            return {ok: false, path: pdf_path, err: "导出命令已执行，但未生成 PDF 文件"}
        if (FileGetSize(pdf_path) <= 0)
            return {ok: false, path: pdf_path, err: "生成的 PDF 文件为空"}
        return {ok: true, path: pdf_path, err: ""}
    } catch Error as err {
        return {ok: false, path: pdf_path, err: err.Message}
    }
}

FormatInlineHeading(doc, p, raw_text, hFont, hSize, hBold, bFont, bSize, bBold, maxLen, keepNext := false, standaloneAlign := 3) {
    total_len := StrLen(raw_text)

    if (total_len <= maxLen) {
        SetRangeFont(p.Range, hFont, hSize, hBold)
        p.Format.Alignment := standaloneAlign
        p.Format.CharacterUnitFirstLineIndent := 2
        if (keepNext)
            p.Format.KeepWithNext := true
        return
    }

    split_pos := RegExMatch(raw_text, "[。！？]")
    if (split_pos > 0 && split_pos <= maxLen) {
        SetRangeFont(doc.Range(p.Range.Start, p.Range.Start + split_pos), hFont, hSize, hBold)
        if (split_pos < total_len)
            SetRangeFont(doc.Range(p.Range.Start + split_pos, p.Range.End), bFont, bSize, bBold)
        p.Format.Alignment := 3
        p.Format.CharacterUnitFirstLineIndent := 2
        if (keepNext)
            p.Format.KeepWithNext := true
    } else {
        SetRangeFont(p.Range, bFont, bSize, bBold)
        p.Format.Alignment := 3
        p.Format.CharacterUnitFirstLineIndent := 2
    }
}

; =====================================================================
; [核心排版引擎] 单文档格式化主流程
; =====================================================================
FormatWordDocument(file_path, app, &err_msg := "", &warn_msg := "") {
    global MarginTop, MarginBottom, MarginLeft, MarginRight, PageNumFormat, SkipSinglePageNum
    global ActiveStyle, LINE_SPACING_RULE, LINE_SPACING_BODY, LINE_SPACING_TITLE
    global ExportPDF, CreateBAK, OptFixDocNum, OptPreventOrphan, BakDir, DefaultBakDir, MaxHeadingLen

    err_msg := "", warn_msg := ""

    if (!FileExist(file_path)) {
        err_msg := "文件不存在或无访问权限"
        return false
    }

    try {
        f := FileOpen(file_path, "rw")
        if (f)
            f.Close()
    } catch {
        err_msg := "文件被占用（可能正被 Word 或其他程序打开）"
        return false
    }

    if (CreateBAK) {
        targetDir := (BakDir != "") ? BakDir : DefaultBakDir
        if (!DirExist(targetDir)) {
            try DirCreate(targetDir)
        }
        SplitPath(file_path, , , &fileExt, &nameNoExt)
        bak_path := targetDir "\" nameNoExt "_" FormatTime(, "yyyyMMdd_HHmmss") "." fileExt
        try FileCopy(file_path, bak_path, true)
    }

    doc := ""
    try {
        try app.ScreenUpdating := false
        doc := app.Documents.Open(file_path, false, false, false)

        ptMTitle   := ParseFontSize(ActiveStyle.MTitleSize, 22.0)
        ptSubTitle := ParseFontSize(ActiveStyle.SubTitleSize, 16.0)
        ptAuthor   := ParseFontSize(ActiveStyle.AuthorSize, 16.0)
        ptBody     := ParseFontSize(ActiveStyle.BodySize, 16.0)
        ptH1       := ParseFontSize(ActiveStyle.H1Size, 16.0)
        ptH2       := ParseFontSize(ActiveStyle.H2Size, 16.0)
        ptH3       := ParseFontSize(ActiveStyle.H3Size, 16.0)
        ptH4       := ParseFontSize(ActiveStyle.H4Size, 16.0)
        ptTable    := ParseFontSize(ActiveStyle.TableSize, 10.5)

        try doc.Content.ListFormat.ConvertNumbersToText()

        if (OptFixDocNum) {
            check_limit := Min(doc.Paragraphs.Count, 5)
            Loop check_limit {
                p_check := doc.Paragraphs(A_Index)
                if (p_check.Range.Tables.Count == 0) {
                    txt_check := TrimAllSpaces(RTrim(p_check.Range.Text, "`r`n"))
                    if RegExMatch(txt_check, "^([^0-9\r\n]{2,15})[\(\[\{（【〔]([12]\d{3})[\)\]\}）>〕]\s*(\d+)\s*号$", &m) {
                        SetParagraphText(p_check, m[1] . "〔" . m[2] . "〕" . m[3] . "号")
                        SetRangeFont(p_check.Range, ActiveStyle.BodyFont, ptBody, ActiveStyle.BodyBold)
                        p_check.Format.Alignment := 1
                        p_check.Format.CharacterUnitFirstLineIndent := 0
                        p_check.Format.FirstLineIndent := 0
                        break
                    }
                }
            }
        }

        page_setup := doc.PageSetup
        page_setup.TopMargin      := MmToPt(MarginTop)
        page_setup.BottomMargin   := MmToPt(MarginBottom)
        page_setup.LeftMargin     := MmToPt(MarginLeft)
        page_setup.RightMargin    := MmToPt(MarginRight)
        page_setup.HeaderDistance := MmToPt(15)
        page_setup.FooterDistance := MmToPt(28)

        CleanAllBlankParagraphs(doc)

        content := doc.Content
        SetRangeFont(content, ActiveStyle.BodyFont, ptBody, ActiveStyle.BodyBold)
        content.ParagraphFormat.Alignment := 3
        content.ParagraphFormat.CharacterUnitLeftIndent := 0
        content.ParagraphFormat.LeftIndent := 0
        content.ParagraphFormat.CharacterUnitRightIndent := 0
        content.ParagraphFormat.RightIndent := 0
        content.ParagraphFormat.CharacterUnitFirstLineIndent := 2
        content.ParagraphFormat.LineSpacingRule := LINE_SPACING_RULE
        content.ParagraphFormat.LineSpacing     := LINE_SPACING_BODY
        content.ParagraphFormat.SpaceBeforeAuto := 0
        content.ParagraphFormat.SpaceAfterAuto  := 0
        content.ParagraphFormat.LineUnitBefore  := 0
        content.ParagraphFormat.LineUnitAfter   := 0
        content.ParagraphFormat.SpaceBefore     := 0
        content.ParagraphFormat.SpaceAfter      := 0
        content.ParagraphFormat.KeepWithNext    := false
        
        Loop doc.Paragraphs.Count {
            if (Mod(A_Index, 20) == 0)
                Sleep(-1)
            p_tmp := doc.Paragraphs(A_Index)
            if (p_tmp.Range.Tables.Count == 0)
                GetCleanText(p_tmp)
        }

        if (doc.Paragraphs.Count == 0) {
            try {
                doc.Save()
                doc.Saved := true
            }
            return true
        }

        Loop doc.Paragraphs.Count {
            if (A_Index > doc.Paragraphs.Count)
                break
            p_curr := doc.Paragraphs(A_Index)
            if (p_curr.Range.Tables.Count > 0)
                continue
            t_curr := TrimAllSpaces(RTrim(p_curr.Range.Text, "`r`n"))

            if (RegExMatch(t_curr, "^附件\s*[:：]")) {
                att_idx := A_Index
                if (att_idx > 1) {
                    p_prev := doc.Paragraphs(att_idx - 1)
                    if (p_prev.Range.Tables.Count == 0) {
                        t_prev := TrimAllSpaces(RTrim(p_prev.Range.Text, "`r`n"))
                        if (t_prev != "") {
                            p_curr.Range.InsertParagraphBefore()
                            p_blank_att := doc.Paragraphs(att_idx)
                            p_blank_att.Format.CharacterUnitFirstLineIndent := 0
                            p_blank_att.Format.FirstLineIndent := 0
                            p_blank_att.Format.LineSpacingRule := LINE_SPACING_RULE
                            p_blank_att.Format.LineSpacing     := LINE_SPACING_BODY
                            att_idx++
                            p_curr := doc.Paragraphs(att_idx)
                        }
                    }
                }

                if (SubStr(t_curr, -1) == "。")
                    t_curr := SubStr(t_curr, 1, StrLen(t_curr) - 1)

                if RegExMatch(t_curr, "^附件\s*[:：]\s*(\d+)\s*[\.．、·](?!\d)\s*(.*)$", &m_att1) {
                    t_curr := "附件：" . m_att1[1] . "." . LTrimAllSpaces(m_att1[2])
                    SetParagraphText(p_curr, t_curr)
                    wrapChars := 5.0 + GetVisualLength(m_att1[1] . ".")
                } else if RegExMatch(t_curr, "^附件\s*[:：]\s*(.*)$", &m_att0) {
                    t_curr := "附件：" . LTrimAllSpaces(m_att0[1])
                    SetParagraphText(p_curr, t_curr)
                    wrapChars := 5.0
                } else {
                    wrapChars := 5.0
                }

                SetRangeFont(p_curr.Range, ActiveStyle.BodyFont, ptBody, ActiveStyle.BodyBold)
                p_curr.Format.CharacterUnitLeftIndent := 0
                p_curr.Format.CharacterUnitFirstLineIndent := 0
                p_curr.Format.LeftIndent := wrapChars * ptBody
                p_curr.Format.FirstLineIndent := (2.0 - wrapChars) * ptBody
                p_curr.Format.Alignment := 0

                next_idx := att_idx + 1
                while (next_idx <= doc.Paragraphs.Count) {
                    p_sub := doc.Paragraphs(next_idx)
                    if (p_sub.Range.Tables.Count > 0)
                        break
                    t_sub := TrimAllSpaces(RTrim(p_sub.Range.Text, "`r`n"))
                    if RegExMatch(t_sub, "^(\d+)\s*[\.．、·](?!\d)\s*(.*)$", &m_sub) {
                        sub_title := LTrimAllSpaces(m_sub[2])
                        if (SubStr(sub_title, -1) == "。")
                            sub_title := SubStr(sub_title, 1, StrLen(sub_title) - 1)

                        num_prefix := m_sub[1] . "."
                        t_sub := num_prefix . sub_title
                        SetParagraphText(p_sub, t_sub)

                        prefixLen := GetVisualLength(num_prefix)
                        SetRangeFont(p_sub.Range, ActiveStyle.BodyFont, ptBody, ActiveStyle.BodyBold)
                        p_sub.Format.CharacterUnitLeftIndent := 0
                        p_sub.Format.CharacterUnitFirstLineIndent := 0
                        p_sub.Format.LeftIndent := (5.0 + prefixLen) * ptBody
                        p_sub.Format.FirstLineIndent := -(prefixLen * ptBody)
                        p_sub.Format.Alignment := 0
                        next_idx++
                    } else {
                        break
                    }
                }
                break
            }
        }

        header_count := 0
        p1 := doc.Paragraphs(1)
        if (p1.Range.Tables.Count == 0) {
            text1 := TrimAllSpaces(RTrim(p1.Range.Text, "`r`n"))
            if (text1 != "" && !IsEndWithPunct(text1) && RegExMatch(text1, "^\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日") == 0) {
                header_count := 1
                SetRangeFont(p1.Range, ActiveStyle.MTitleFont, ptMTitle, ActiveStyle.MTitleBold)
                p1.Format.Alignment := 1
                p1.Format.CharacterUnitFirstLineIndent := 0
                p1.Format.FirstLineIndent := 0
                p1.Format.LineSpacing := LINE_SPACING_TITLE

                Loop 2 {
                    idx := A_Index + 1
                    if (idx <= doc.Paragraphs.Count) {
                        p_next := doc.Paragraphs(idx)
                        if (p_next.Range.Tables.Count > 0)
                            break
                        next_text := TrimAllSpaces(RTrim(p_next.Range.Text, "`r`n"))
                        if (RegExMatch(next_text, "^\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日") > 0)
                            break
                        if (next_text != "" && !IsEndWithPunct(next_text)) {
                            header_count := idx
                            isSub := (StrLen(next_text) > 5)
                            SetRangeFont(p_next.Range, isSub ? ActiveStyle.SubTitleFont : ActiveStyle.AuthorFont, isSub ? ptSubTitle : ptAuthor, isSub ? ActiveStyle.SubTitleBold : ActiveStyle.AuthorBold)
                            p_next.Format.Alignment := 1
                            p_next.Format.CharacterUnitFirstLineIndent := 0
                            p_next.Format.FirstLineIndent := 0
                        } else {
                            break
                        }
                    }
                }
            }
        }

        start_body_idx := 1
        if (header_count > 0) {
            r := doc.Paragraphs(header_count).Range
            r.Collapse(0)
            r.InsertParagraphAfter()

            blank_p := doc.Paragraphs(header_count + 1)
            blank_p.Format.CharacterUnitFirstLineIndent := 0
            blank_p.Format.FirstLineIndent := 0
            blank_p.Format.LineSpacingRule := LINE_SPACING_RULE
            blank_p.Format.LineSpacing     := LINE_SPACING_BODY
            start_body_idx := header_count + 2
        }

        if (start_body_idx <= doc.Paragraphs.Count) {
            body_range := doc.Range(doc.Paragraphs(start_body_idx).Range.Start, doc.Content.End)
            find_obj := body_range.Find
            find_obj.ClearFormatting()
            find_obj.Replacement.ClearFormatting()
            find_obj.Wrap := 0
            try find_obj.Execute("^l", , , , , , , 0, , "^p", 2)
        }

        next_is_attachment_title := false
        p_idx := start_body_idx

        while (p_idx <= doc.Paragraphs.Count) {
            if (Mod(p_idx, 15) == 0)
                Sleep(-1)

            p := doc.Paragraphs(p_idx)
            if (p.Range.Tables.Count > 0) {
                p_idx++
                continue
            }

            raw_text := RTrim(p.Range.Text, "`r`n")
            text := TrimAllSpaces(StrReplace(raw_text, Chr(12), ""))
            if (text == "") {
                p_idx++
                continue
            }

            if (RegExMatch(text, "^附件\s*\d*$")) {
                SetRangeFont(p.Range, ActiveStyle.BodyFont, ptBody, ActiveStyle.BodyBold)
                p.Format.Alignment := 0
                p.Format.CharacterUnitFirstLineIndent := 0
                p.Format.FirstLineIndent := 0
                p.Format.LineSpacingRule := LINE_SPACING_RULE
                p.Format.LineSpacing     := LINE_SPACING_BODY
                next_is_attachment_title := true
                p_idx++
                continue
            }

            if (next_is_attachment_title) {
                SetRangeFont(p.Range, ActiveStyle.MTitleFont, ptMTitle, ActiveStyle.MTitleBold)
                p.Format.Alignment := 1
                p.Format.CharacterUnitFirstLineIndent := 0
                p.Format.FirstLineIndent := 0
                p.Format.LineSpacingRule := 4
                p.Format.LineSpacing     := LINE_SPACING_TITLE
                next_is_attachment_title := false

                r_att := p.Range
                r_att.Collapse(0)
                r_att.InsertParagraphAfter()

                p_blank := doc.Paragraphs(p_idx + 1)
                p_blank.Format.CharacterUnitFirstLineIndent := 0
                p_blank.Format.FirstLineIndent := 0
                p_blank.Format.LineSpacingRule := LINE_SPACING_RULE
                p_blank.Format.LineSpacing     := LINE_SPACING_BODY

                p_idx += 2
                continue
            }

            if (SubStr(text, 1, 3) == "【引】") {
                clean_quote := LTrimAllSpaces(SubStr(text, 4))
                SetParagraphText(p, clean_quote)
                SetRangeFont(p.Range, ActiveStyle.H2Font, ptH2, ActiveStyle.H2Bold)
                p.Format.Alignment := 3
                p.Format.CharacterUnitFirstLineIndent := 2
                p.Format.LineSpacingRule := LINE_SPACING_RULE
                p.Format.LineSpacing     := LINE_SPACING_BODY
                p_idx++
                continue
            }

            if (RegExMatch(text, "^附件\s*[:：]")) {
                p_idx++
                while (p_idx <= doc.Paragraphs.Count) {
                    p_sub := doc.Paragraphs(p_idx)
                    if (p_sub.Range.Tables.Count > 0)
                        break
                    t_sub := TrimAllSpaces(RTrim(p_sub.Range.Text, "`r`n"))
                    if (RegExMatch(t_sub, "^\d+\s*[\.．、·](?!\d)"))
                        p_idx++
                    else
                        break
                }
                continue
            }

            if (p_idx == start_body_idx && IsEndWithColon(text)) {
                p.Format.Alignment := 0
                p.Format.CharacterUnitFirstLineIndent := 0
                p.Format.FirstLineIndent := 0
                p_idx++
                continue
            }

            if (RegExMatch(text, "^([一二三四五六七八九十百]+)\s*([、.．·])\s*(.*)$", &m_h1) && !RegExMatch(m_h1[3], "^[一二三四五六七八九十百]+[、年]")) {
                clean_h1 := m_h1[1] . "、" . LTrimAllSpaces(m_h1[3])
                if (raw_text != clean_h1) {
                    SetParagraphText(p, clean_h1)
                    raw_text := clean_h1
                    text := clean_h1
                }
                FormatInlineHeading(doc, p, raw_text, ActiveStyle.H1Font, ptH1, ActiveStyle.H1Bold, ActiveStyle.BodyFont, ptBody, ActiveStyle.BodyBold, MaxHeadingLen, OptPreventOrphan, 0)
            } else if RegExMatch(text, "^(\x28|（)\s*([一二三四五六七八九十百]+)\s*(\x29|）)\s*(.*)$", &m_h2) {
                clean_h2 := m_h2[1] . m_h2[2] . m_h2[3] . LTrimAllSpaces(m_h2[4])
                if (raw_text != clean_h2) {
                    SetParagraphText(p, clean_h2)
                    raw_text := clean_h2
                    text := clean_h2
                }
                FormatInlineHeading(doc, p, raw_text, ActiveStyle.H2Font, ptH2, ActiveStyle.H2Bold, ActiveStyle.BodyFont, ptBody, ActiveStyle.BodyBold, MaxHeadingLen, OptPreventOrphan, 3)
            } else if RegExMatch(text, "^(\d+)\s*([.．、·])(?!\d)\s*(.*)$", &m_h3) {
                clean_h3 := m_h3[1] . "." . LTrimAllSpaces(m_h3[3])
                if (raw_text != clean_h3) {
                    SetParagraphText(p, clean_h3)
                    raw_text := clean_h3
                    text := clean_h3
                }
                FormatInlineHeading(doc, p, raw_text, ActiveStyle.H3Font, ptH3, ActiveStyle.H3Bold, ActiveStyle.BodyFont, ptBody, ActiveStyle.BodyBold, MaxHeadingLen, OptPreventOrphan, 3)
            } else if RegExMatch(text, "^(\x28|（)\s*(\d+)\s*(\x29|）)\s*(.*)$", &m_h4) {
                clean_h4 := m_h4[1] . m_h4[2] . m_h4[3] . LTrimAllSpaces(m_h4[4])
                if (raw_text != clean_h4) {
                    SetParagraphText(p, clean_h4)
                    raw_text := clean_h4
                    text := clean_h4
                }
                FormatInlineHeading(doc, p, raw_text, ActiveStyle.H4Font, ptH4, ActiveStyle.H4Bold, ActiveStyle.BodyFont, ptBody, ActiveStyle.BodyBold, MaxHeadingLen, OptPreventOrphan, 3)
            } else {
                SetRangeFont(p.Range, ActiveStyle.BodyFont, ptBody, ActiveStyle.BodyBold)
                p.Format.Alignment := 3
                p.Format.CharacterUnitFirstLineIndent := 2
                p.Format.LineSpacingRule := LINE_SPACING_RULE
                p.Format.LineSpacing     := LINE_SPACING_BODY
            }

            p_idx++
        }

        date_idx := 0, signer_idx := 0
        final_total := doc.Paragraphs.Count

        Loop final_total {
            idx := final_total - A_Index + 1
            p_chk := doc.Paragraphs(idx)
            if (p_chk.Range.Tables.Count > 0)
                continue
            txt := TrimAllSpaces(RTrim(p_chk.Range.Text, "`r`n"))
            if (RegExMatch(txt, "^\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日") > 0) {
                date_idx := idx
                break
            }
        }

        if (date_idx > 0) {
            while (date_idx > 1) {
                p_above := doc.Paragraphs(date_idx - 1)
                if (p_above.Range.Tables.Count > 0)
                    break
                above_txt := TrimAllSpaces(RTrim(p_above.Range.Text, "`r`n"))
                if (RegExMatch(above_txt, "^\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日") > 0) {
                    try {
                        p_above.Range.Delete()
                        date_idx--
                    } catch {
                        break
                    }
                } else {
                    break
                }
            }

            cand_idx := date_idx - 1
            while (cand_idx >= 1) {
                p_cand := doc.Paragraphs(cand_idx)
                if (p_cand.Range.Tables.Count > 0) {
                    cand_idx--
                    continue
                }
                cand_txt := TrimAllSpaces(RTrim(p_cand.Range.Text, "`r`n"))
                if (cand_txt == "") {
                    cand_idx--
                    continue
                }
                isDateLine := (RegExMatch(cand_txt, "^\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日") > 0)
                if (!IsEndWithPunct(cand_txt) && !RegExMatch(cand_txt, "^(\d+[\.．、]|附件)") && !isDateLine && StrLen(cand_txt) <= 25)
                    signer_idx := cand_idx
                break
            }

            sign_top_idx := (signer_idx > 0) ? signer_idx : date_idx
            base_pages := doc.ComputeStatistics(2)

            doc.Paragraphs(sign_top_idx).Range.InsertParagraphBefore()
            p_blank1 := doc.Paragraphs(sign_top_idx)
            p_blank1.Format.CharacterUnitFirstLineIndent := 0
            p_blank1.Format.FirstLineIndent := 0
            p_blank1.Format.LineSpacingRule := LINE_SPACING_RULE
            p_blank1.Format.LineSpacing     := LINE_SPACING_BODY

            doc.Paragraphs(sign_top_idx + 1).Range.InsertParagraphBefore()
            p_blank2 := doc.Paragraphs(sign_top_idx + 1)
            p_blank2.Format.CharacterUnitFirstLineIndent := 0
            p_blank2.Format.FirstLineIndent := 0
            p_blank2.Format.LineSpacingRule := LINE_SPACING_RULE
            p_blank2.Format.LineSpacing     := LINE_SPACING_BODY

            if (doc.ComputeStatistics(2) > base_pages) {
                doc.Paragraphs(sign_top_idx + 1).Range.Delete()
                inserted_count := 1
            } else {
                inserted_count := 2
            }

            if (signer_idx > 0)
                signer_idx += inserted_count
            date_idx += inserted_count

            p_date := doc.Paragraphs(date_idx)
            clean_date := RegExReplace(TrimAllSpaces(RTrim(p_date.Range.Text, "`r`n")), "\s+", " ")

            clean_signer := ""
            if (signer_idx > 0) {
                p_signer := doc.Paragraphs(signer_idx)
                clean_signer := TrimAllSpaces(RTrim(p_signer.Range.Text, "`r`n"))
            }

            if (signer_idx > 0 && clean_signer != "") {
                lenSigner := GetVisualLength(clean_signer)
                lenDate   := GetVisualLength(clean_date)

                SetParagraphText(p_signer, clean_signer)
                SetParagraphText(p_date, clean_date)

                SetRangeFont(p_signer.Range, ActiveStyle.BodyFont, ptBody, ActiveStyle.BodyBold)
                SetRangeFont(p_date.Range, ActiveStyle.BodyFont, ptBody, ActiveStyle.BodyBold)

                if (lenSigner > lenDate) {
                    signerRightChars := 2.0
                    dateRightChars   := Round(2.0 + (lenSigner - lenDate) / 2.0, 1)
                } else {
                    dateRightChars   := 4.0
                    signerRightChars := Round(4.0 + (lenDate - lenSigner) / 2.0, 1)
                }

                p_signer.Format.FirstLineIndent := 0
                p_signer.Format.CharacterUnitFirstLineIndent := 0
                p_signer.Format.LeftIndent := 0
                p_signer.Format.CharacterUnitLeftIndent := 0
                p_signer.Format.RightIndent := 0
                p_signer.Format.CharacterUnitRightIndent := signerRightChars
                p_signer.Format.Alignment := 2

                p_date.Format.FirstLineIndent := 0
                p_date.Format.CharacterUnitFirstLineIndent := 0
                p_date.Format.LeftIndent := 0
                p_date.Format.CharacterUnitLeftIndent := 0
                p_date.Format.RightIndent := 0
                p_date.Format.CharacterUnitRightIndent := dateRightChars
                p_date.Format.Alignment := 2
            } else {
                SetParagraphText(p_date, clean_date)
                SetRangeFont(p_date.Range, ActiveStyle.BodyFont, ptBody, ActiveStyle.BodyBold)
                p_date.Format.FirstLineIndent := 0
                p_date.Format.CharacterUnitFirstLineIndent := 0
                p_date.Format.LeftIndent := 0
                p_date.Format.CharacterUnitLeftIndent := 0
                p_date.Format.RightIndent := 0
                p_date.Format.CharacterUnitRightIndent := 4.0
                p_date.Format.Alignment := 2
            }
        }

        RemoveTrailingEmptyParagraphs(doc)

        ; -----------------------------------------------------------------
        ; 步骤 8：后置处理表格字体、字号、段落对齐与缩进（默认强制生效，单元格文字左对齐）
        ; -----------------------------------------------------------------
        if (doc.Tables.Count > 0) {
            for tbl in doc.Tables {
                try {
                    tblRange := tbl.Range
                    SetRangeFont(tblRange, ActiveStyle.TableFont, ptTable, ActiveStyle.TableBold)
                    tblParaFormat := tblRange.ParagraphFormat
                    tblParaFormat.CharacterUnitFirstLineIndent := 0
                    tblParaFormat.FirstLineIndent := 0
                    tblParaFormat.CharacterUnitLeftIndent := 0
                    tblParaFormat.LeftIndent := 0
                    tblParaFormat.LineSpacingRule := 0
                    tblParaFormat.SpaceBefore     := 0
                    tblParaFormat.SpaceAfter      := 0
                    tbl.Rows.Alignment := 1 ; 1 = wdAlignRowCenter（表格整体居中）
                    for cell in tbl.Range.Cells {
                        cell.VerticalAlignment := 1 ; 1 = wdCellAlignVerticalCenter（单元格垂直居中）
                        for cellPara in cell.Paragraphs {
                            cellPara.Format.Alignment := 0 ; 0 = wdAlignParagraphLeft（单元格文字左对齐）
                        }
                    }
                }
            }
        }

        pageNumWarn := SetPageNumbers(doc, PageNumFormat, SkipSinglePageNum)
        if (pageNumWarn != "")
            warn_msg := pageNumWarn

        try {
            doc.Save()
            doc.Saved := true
        } catch {
            err_msg := "保存失败，文件可能在排版中途被锁定或无权限"
            throw Error(err_msg)
        }

        if (ExportPDF) {
            pdfResult := ExportToPdf(doc, file_path)
            if (!pdfResult.ok) {
                pdfWarn := "PDF 导出失败（" pdfResult.path "）：" pdfResult.err
                warn_msg := (warn_msg == "") ? pdfWarn : warn_msg "；" pdfWarn
            }
        }

        try app.ScreenUpdating := true
        return true

    } catch Error as err {
        try app.ScreenUpdating := true
        err_msg := err.Message
        return false
    } finally {
        if (IsSet(doc) && doc != "") {
            try {
                doc.Saved := true
                doc.Close(0)
            }
        }
    }
}

; =====================================================================
; [高级设置：标题拦截阈值子窗口]
; =====================================================================
OpenHeadingLenSettings(*) {
    global mainGui, MaxHeadingLen
    lenGui := Gui("+Owner" mainGui.Hwnd " +ToolWindow", "标题字数拦截设置")
    lenGui.SetFont("s9", "Microsoft YaHei")
    lenGui.Add("Text", "x20 y20 w280", "排版时，如果带有标题序号的段落字数超过此阈值，`n将被强制视为【正文】而不予加粗（防止长段落误判）。")
    
    lenGui.SetFont("s10")
    lenGui.Add("Text", "x50 y73 w70", "最大字数:")
    edLen := lenGui.Add("Edit", "x125 y70 w65 Center Number", MaxHeadingLen)
    lenGui.Add("UpDown", "Range10-200", MaxHeadingLen)
    lenGui.Add("Text", "x+10 yp+3", "字 (默认 35)")
    
    btnSave := lenGui.Add("Button", "x110 y120 w100 h30 Default", "保存设置")
    btnSave.OnEvent("Click", (*) => SaveHeadingLen())
    
    SaveHeadingLen() {
        val := edLen.Value
        if (val != "" && IsNumber(val) && val > 0) {
            MaxHeadingLen := Integer(val)
            SetConfigVal("Options", "MaxHeadingLen", MaxHeadingLen)
            lenGui.Destroy()
            ShowFloatingToast("设置已保存", "标题拦截字数已更改为 " MaxHeadingLen " 字", 2000, true)
        }
    }
    lenGui.Show("w320 h170")
}

; =====================================================================
; [帮助菜单：关于本程序 / 打赏作者弹窗]
; =====================================================================
ShowAboutDialog(*) {
    global mainGui, APP_VERSION
    abtGui := Gui("+Owner" mainGui.Hwnd " +ToolWindow", "请作者喝一杯咖啡")
    abtGui.SetFont("s14 bold c0055AA", "Microsoft YaHei")
    abtGui.Add("Text", "x20 y20 w410 Center", "公文自动排版引擎")
    
    abtGui.SetFont("s10 norm c000000")
    infoText := "版本：V" APP_VERSION "`n(国标 GB/T 9704-2012)`n`n"
             . "作者：Eyes`n"
             . "ＱＱ：28525626`n`n"
             . "如果这个小工具帮到了您，`n欢迎请作者喝杯咖啡支持一下！"
    abtGui.Add("Text", "x30 y75 w220", infoText)
    
    qrBase64 := "
    (
    
    )"
    
    qrCodePath := Base64ToTempFile(qrBase64)
    if (qrCodePath != "" && FileExist(qrCodePath)) {
        abtGui.Add("Picture", "x260 y70 w160 h-1", qrCodePath)
    } else {
        abtGui.SetFont("s9 c777777")
        abtGui.Add("Text", "x260 y140 w160 Center", "(收款码加载失败)")
        abtGui.SetFont("s10 c000000")
    }
    
    btnOK := abtGui.Add("Button", "x175 y280 w100 h32 Default", "确 定")
    CleanTempFile(*) {
        abtGui.Destroy()
        if (qrCodePath != "" && FileExist(qrCodePath)) {
            try FileDelete(qrCodePath)
        }
    }
    btnOK.OnEvent("Click", CleanTempFile)
    abtGui.OnEvent("Close", CleanTempFile)
    abtGui.Show("w450 h330")
}

Base64ToTempFile(base64Str) {
    tempPath := A_Temp "\DocFormatTool_QRCode_" A_TickCount ".jpg"
    base64Str := RegExReplace(base64Str, "\s") 
    if (base64Str == "")
        return ""
    size := 0
    DllCall("crypt32\CryptStringToBinary", "Str", base64Str, "UInt", 0, "UInt", 1, "Ptr", 0, "Ptr", 0)
    if (size == 0)
        return ""
    buf := Buffer(size)
    DllCall("crypt32\CryptStringToBinary", "Str", base64Str, "UInt", 0, "UInt", 1, "Ptr", buf, "Ptr", 0, "Ptr", 0)
    try {
        f := FileOpen(tempPath, "w")
        f.RawWrite(buf, size)
        f.Close()
        return tempPath
    } catch {
        return ""
    }
}