
CoordMode("Mouse", "Window")
SendMode("Input")
#SingleInstance Force
; persistent 1
#WinActivateForce
SetTitleMatchMode(2)
SetControlDelay(1)
SetWinDelay(0)
SetKeyDelay(-1, -1)
SetMouseDelay(-1)

#Include "Yaml.ahk"

copynow(all := 0, timeout := 1) {
	svclp := A_Clipboard
	A_Clipboard := ""
	Send((all ? "^a" : "") "^c")
	; Sleep(50)
	Errorlevel := !ClipWait(timeout, 1)
	result := A_Clipboard
	A_Clipboard := svclp
	return result
}

failedmsg(Text, condition := 1) {
	if (condition) {
		msgbox(Text "`nExiting", "Error", 48)
		Exit()
	}
}

FileOverwrite(Text,Filename){
	try FileDelete Filename
	FileAppend Text,Filename
}

ThousandsSep(x, s := ",") => RegExReplace(x, "\G\d+?(?=(\d{3})+(?:\D|$))", "$0" s)

