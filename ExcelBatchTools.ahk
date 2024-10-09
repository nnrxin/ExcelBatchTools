;基础参数设置
KeyHistory 0
ListLines 0
SendMode "Input"
#NoTrayIcon               ;无托盘图标
#SingleInstance Ignore    ;不能双开
#Include <_BasicLibs_>
#Include <File\Path>
#Include <File\FileCopyEx>
#Include <GUI\ProgressGui>
#Include <GUI\ProgressInStatusBar>


; APP名称
APP_NAME      := "EBT"
;@Ahk2Exe-Let U_NameShort = %A_PriorLine~U)(^.*")|(".*$)%
; APP全称
APP_NAME_FULL := "ExcelBatchTools"
;@Ahk2Exe-Let U_Name = %A_PriorLine~U)(^.*")|(".*$)%
; APP中文名称
APP_NAME_CN   := "EXCEL文件批量处理工具EBT"
;@Ahk2Exe-Let U_NameCN = %A_PriorLine~U)(^.*")|(".*$)%
; 当前版本
APP_VERSION   := "0.0.2"
;@Ahk2Exe-Let U_ProductVersion = %A_PriorLine~U)(^.*")|(".*$)%


;编译后文件名
;@Ahk2Exe-Obey U_bits, = %A_PtrSize% * 8
;@Ahk2Exe-ExeName %U_NameCN%(%U_bits%bit) v%U_ProductVersion%
;编译后属性信息
;@Ahk2Exe-SetName %U_Name%
;@Ahk2Exe-SetProductVersion %U_ProductVersion%
;@Ahk2Exe-SetLanguage 0x0804
;@Ahk2Exe-SetCopyright Copyright (c) 2024 nnrxin
;编译后的图标(与脚本名同目录同名的ico文件,不存在时会报错)
;@Ahk2Exe-SetMainIcon %A_ScriptName~\.[^\.]+$~.ico%



AHK_DATA_DIR_PATH := A_AppData "\AHKDATA"

;安装XL\XL库文件
#Include AHK_InstallFiles.ahk
if !AHK_DirInstallTo(AHK_DATA_DIR_PATH)    ;非覆盖安装
	MsgBox "文件安装错误!"
DllCall('LoadLibrary', 'str', AHK_DATA_DIR_PATH '\XL\' (A_PtrSize * 8) 'bit\libxl.dll', 'ptr')
#Include <XL\XL>


;APP保存信息(存储在AppData)
APP_DATA_PATH := A_AppData "\" APP_NAME_FULL                    ;在系统AppData的保存位置
APP_DATA_CACHE_PATH := APP_DATA_PATH "\cache"                   ;缓存文件路径
DirCreate APP_DATA_CACHE_PATH                                   ;路径不存在时需要新建
APP_INI := IniSaved(APP_DATA_PATH "\" APP_NAME "_config.ini")   ;创建配置ini类


;全局参数
G := {}

;=================================
;↓↓↓↓↓↓↓↓↓  MainGUI 构建 ↓↓↓↓↓↓↓↓↓
;=================================

;创建主GUI
MainGuiWidth := 700, MainGuiHeight := 500
MainGui := Gui("+Resize +MinSize" MainGuiWidth "x" MainGuiHeight , APP_NAME_CN " " APP_VERSION)   ;GUI可修改尺寸
MainGui.Show("hide w" MainGuiWidth " h" MainGuiHeight)
MainGui.MarginX := MainGui.MarginY := 0
MainGui.SetFont("s9", "微软雅黑")
;MainGui.BackColor := 0xCCE8CF   ;护眼蓝色
GroupWidth := 125 ; 右侧框架宽度

;增加Guitooltip
MainGui.Tips := GuiCtrlTips(MainGui)

;列表框
LV := MainGui.Add("ListView", "Section xm+5 ym+5 w" MainGuiWidth-GroupWidth-17 " h" MainGuiHeight-30 " AW AH", ["文件名","路径","大小","状态"])
LV.ModifyCol(3, "Right"), LV.ModifyCol(4, "Center")
;列表加载文件
filesInLV := Map()
LV.LoadFilesAndDirs := LV_LoadFilesAndDirs
LV_LoadFilesAndDirs(this, pathArray) {
	static exts := ["xls","xlsx"]
	this.Opt("-Redraw")
	files := []
	for _, path in pathArray {
		if DirExist(path) {
			Loop Files, path "\*.*", "FR"
				files.Push({path: A_LoopFileFullPath, midPath: Path_Relative(A_LoopFileFullPath, Path_Dir(path))})
			continue
		}
		files.Push({path: path, midPath: ""})
	}
	for i, file in files {
		if filesInLV.Has(file.path)
			continue
		SplitPath file.path, &OutFileName, &OutDir, &OutExtension, &OutNameNoExt, &OutDrive
		if !exts.IndexOf(OutExtension)
		or InStr(FileGetAttrib(file.path), "H") ; 跳过隐藏文件
        or FileGetSize(file.path, "KB") < 1
			continue
		f := filesInLV[file.path] := {path: file.path, name: OutFileName, sizeKB: Format("{:.1f} KB", FileGetSize(file.path)/1024), status: "等待处理", midPath: file.midPath}
		this.Add("Icon" this.LoadFileIcon(f.path), f.name, f.path, f.sizeKB, f.status)
	}
	this.AdjustColumnsWidth()
	this.Opt("+Redraw")
	EnableBottons(LV.GetCount()) ; 控制按钮
	SB.SetText("文件总数: " LV.GetCount())
}


;保存方式 Group
MainGui.SetFont("c0070DE bold", "微软雅黑")
MainGui.Add("GroupBox", "Section x+7 ym w" GroupWidth " h70 AX", "保存方式")
MainGui.SetFont("cDefault norm", "微软雅黑")
RD1 := MainGui.Add("Radio", "xp+10 yp+22 AXP Group", "覆盖源文件")
RD1.Value := APP_INI.Init(RD1, "save", "RD1", 0)
RD2 := MainGui.Add("Radio", "xp y+5 AXP", "保存为新文件")
RD2.Value := APP_INI.Init(RD2, "save", "RD2", 1)



;执行 Group
MainGui.SetFont("c0070DE bold", "微软雅黑")
MainGui.Add("GroupBox", "Section xs y+12 w" GroupWidth " h133 AXP", "执行")
MainGui.SetFont("cDefault norm", "微软雅黑")

BTclear := MainGui.Add("Button", "xs+10 yp+22 h27 w105 AXP", "移除所有项")
BTclear.OnEvent("Click", BTclear_Click)
BTclear_Click(thisCtrl, info) {
	LV.Opt("-Redraw")
	LV.Delete()
	filesInLV.Clear()
	LV.Opt("+Redraw")
	EnableBottons(LV.GetCount()) ; 控制按钮
	SB.SetText("移除了所有项")
}

BTremoveFinished := MainGui.Add("Button", "xp y+5 hp wp AXP", "移除成功项")
BTremoveFinished.OnEvent("Click", BTremoveFinished_Click)
BTremoveFinished_Click(thisCtrl, info) {
	LV.Opt("-Redraw")
	deleteRows := []
	Loop LV.GetCount() {
		if LV.GetText(A_Index, 4) != "处理成功"
			continue
		filesInLV.Delete(LV.GetText(A_Index, 2))
		deleteRows.Push(A_Index)
	}
	loop deleteRows.Length
		LV.Delete(deleteRows.Pop())
	LV.Opt("+Redraw")
	EnableBottons(LV.GetCount()) ; 控制按钮
	SB.SetText("移除了完成项")
}

BTstart := MainGui.Add("Button", "xp y+5 h40 wp AXP", "删除前几行")
BTstart.OnEvent("Click", BTstart_Click)
BTstart_Click(thisCtrl, info) {
	;询问删除几行
	IB := InputBox("想要删除EXCEL表格前几行?", "请输入数字")
	if IB.Result = "Cancel"
		return
	n := IB.Value && IsInteger(IB.Value) ? IB.Value : 1
	;覆盖原文件时执行前提醒
	MainGui.Opt("+OwnDialogs")
	if RD1.Value and MsgBox("将在原EXCEL文件上进行修改,是否继续？",, 68) = "No"
		return
	;设置进度条
	MaxCount := LV.GetCount()
	SB.SetText("处理中")
	SB.SetParts(100,100)
	if !SB.HasProp("Progress")
		SB.Progress := ProgressInStatusBar(SB, 0, 3)
	SB.Progress.Value := 0
	SB.Progress.Range := "0-" MaxCount
	SB.Progress.Visible := true
	;开始处理
	EnableBottons(false) ; 禁用按钮
	dirName := APP_NAME_FULL "_" A_Now
	loop MaxCount {
		SB.SetText(A_index "/" MaxCount, 2)
	    file := filesInLV[LV.GetText(A_Index, 2)] ; 确认目标文件名
		try {
			if RD1.Value ; 覆盖原文件
				tarPath := file.path
			else { ; 复制为新文件
				tarPath := A_ScriptDir "\" dirName "\" (file.midPath || file.name)
				DirCreate Path_Dir(tarPath, false)
				FileCopy(file.path, tarPath)
			}
			DeleteTopRows(tarPath, n)
		} catch
		    file.status := "处理失败"
		else
			file.status := "处理成功"
		LV.Modify(A_Index, "Vis Focus Col4", file.status) ; 可见 焦点 选中 列4修改
		SB.Progress.Value++
	}
	LV.AdjustColumnsWidth()
	EnableBottons(true) ; 启用按钮
	SB.Progress.Visible := false
	SB.SetParts()
	SB.SetText("处理完成!")
	if !RD1.Value ; 另存模式时完成后打开新目录
		Run A_ScriptDir "\" dirName "\"
}
;删除最顶层的几行
DeleteTopRows(path, n) {
	book := XL.Load(path), sheet := book[0]
	sheet.removeRow(0, n-1)
	book.save()
	book := ''
}
;启用/禁用按钮函数
EnableBottons(condition) {
	BTstart.Enabled := BTclear.Enabled := BTremoveFinished.Enabled := condition ? true : false
}




;状态栏
SB := MainGui.Add("StatusBar",, "")
SB.SetFont("bold")
SB.SetText("将Excel文件或文件夹拖入窗口中")
;带自动清空的状态栏文字函数
;SB.SetTextWithAutoEmpty(newText, second, partNumber)


;GUI菜单
MainGui.OnEvent("ContextMenu", MainGui_ContextMenu)
MainGui_ContextMenu(GuiObj, GuiCtrlObj, Item, IsRightClick, X, Y) {
	;右键某控件上
	if IsRightClick and GuiCtrlObj and GuiCtrlObj.HasMethod("ContextMenu")
		GuiCtrlObj.ContextMenu(Item, X, Y)
}

;GUI文件拖放
MainGui.OnEvent("DropFiles", MainGui_DropFiles)
MainGui_DropFiles(GuiObj, GuiCtrlObj, FileArray, X, Y) {
	MainGui.Opt("+Disabled")
	;主界面上的拖动
	Switch GuiCtrlObj {
		case LV:
			LV.LoadFilesAndDirs(FileArray)
	}

	MainGui.Opt("-Disabled")
}

;改变GUI尺寸时调整控件
MainGui.OnEvent("Size", MainGui_Size)
MainGui_Size(thisGui, MinMax, W, H) {
	LV.AdjustColumnsWidth()
	if SB.HasProp("Progress")
		SB.Progress.AdjustSize()
}

;GUI关闭
MainGui.OnEvent("Close", MainGui_Close)
MainGui_Close(*) {
	ExitApp
}

;退出APP前运行
OnExit DoBeforeExit
DoBeforeExit(*) {
	MainGui.Hide()
	APP_INI.SaveAll()            ;用户配置保存到ini文件
}


;Gui初始化
DropPaths := Path_InArgs()
if DropPaths.Length
	LV.LoadFilesAndDirs(DropPaths) ; 拖拽文件到程序图标上启动
EnableBottons(LV.GetCount())       ; 按钮初始化


;GUI显示
MainGui.Show("Center")

;=========================
return    ;自动运行段结束 |
;=========================

