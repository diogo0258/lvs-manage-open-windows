
/*
original script posted on <https://autohotkey.com/board/topic/91918-list-open-windows-to-quickly-select-and-close-them/>

Lists visible windows in a searchable listview, from where you can activate or close them.

Press:
- Ctrl+c to select all matching the class of first selected
- Ctrl+p to select all matching the process of first selected
- Ctrl+d to close all selected
- Enter to activate all selected
*/


#NoEnv

	DetectHiddenWindows, Off
	WinGet, AllWinsHwnd, List
	
	AllWinsInfoList := ""
	Loop, % AllWinsHwnd
	{
		CurrHwnd := AllWinsHwnd%A_Index%
		if not IsWindow(CurrHwnd)
			continue
		
		WinGetTitle, CurrTitle, % "ahk_id " CurrHwnd
		StringReplace, CurrTitle, CurrTitle, |, _  ; in order to not clash with delimiter for LVS
		
		WinGet, CurrProc, ProcessName, % "ahk_id " CurrHwnd
		WinGetClass, CurrClass, % "ahk_id " CurrHwnd
		CurrPath := (CurrClass = "CabinetWClass") ? GetFolderFromExplorerWin(CurrHwnd) : ""
		
		AllWinsInfoList .= CurrHwnd "|" CurrTitle "|" CurrClass "|" CurrProc "|" CurrPath "`n"
	}

	LVS_Hwdn := LVS_Init("Callback", "Hwnd|Title|Class|Process|Folder", 1, -1, True, True, 1000)
	GroupAdd, LVS, ahk_id %LVS_Hwdn%
	
	LVS_SetList(AllWinsInfoList, "|")
	LVS_UpdateColOptions("0|300|150|150|AutoHdr")
	LVS_SetBottomText("Enter to activate first selected, ^c/^p to select all with same class/process of first selected, ^d to close all selected")
	LVS_Show()

return


#IfWinActive, ahk_group LVS
	^c::SelectAllWithSameCol(3)
	^p::SelectAllWithSameCol(4)
	^d::callback(LVS_Selected(), False, "close")
#IfWinActive


callback(selectedlist, escaped := False, cmd := "activate") {
	if (escaped or selectedlist = "")
		exitapp
	
	loop, parse, selectedlist, `n
	{
		if (cmd = "activate")
			winactivate, ahk_id %A_LoopField%
		else if (cmd = "close")
			winclose, ahk_id %A_LoopField%
	}
	
	exitapp
}



SelectAllByColValue(col, value) {
	Loop, % LV_GetCount()
	{
		CurrentRow := A_Index
		
		LV_GetText(CurrentValue, CurrentRow, col)
		
		if (CurrentValue = value)
			LV_Modify(CurrentRow, "Select")
	}		
}


SelectAllWithSameCol(col) {  ; searches based on first selected item
	LV_GetText(CurrentValue, LV_GetNext(), col)

	SelectAllByColValue(col, CurrentValue)
}



GetFolderFromExplorerWin(Hwnd) {
	static sa := ComObjCreate("Shell.Application")
	static wins := sa.Windows
	
	loop % wins.Count
	{
		window := wins.Item(A_Index-1)
		if (window.Hwnd = Hwnd)
			break
	}

	path := window.Document.Folder.Self.Path  ; from https://autohotkey.com/board/topic/121208-windows-explorer-get-folder-path/
	
	return path
}


IsWindow(hwnd) ; ManaUser's, http://www.autohotkey.com/forum/viewtopic.php?t=27797
{
	WinGet, s, Style, ahk_id %hwnd%
	return s & 0xC00000 ? (s & 0x80000000 ? 0 : 1) : 0
	;WS_CAPTION AND !WS_POPUP(for tooltips etc) 
}


#Include %A_ScriptDir%\LVS.ahk