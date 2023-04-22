
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

#Include "Misc.ahk"
#Include "UIA.ahk"
#Include "UIA_Browser.ahk"
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

savemsg(Text, Prompt := "") {
	if MsgBox(prompt Text, "Save?", 4) = "Yes"
		A_Clipboard := Text
	else return 0
}

ModMsg(Text := "", Title := "", Options := 0, Buttons := []) {
	switch buttons.Length {
		case 1:
		case 2: Options += 4
		case 3: Options += 3
		default: throw "Unsupported ModMsg"
	}
	SetTimer changebuts, -30
	changebuts() {
		WinWaitActive Title " Modded Msgbox ahk_class #32770"
		for i, but in buttons
			ControlSetText but, "Button" i
	}
	switch MsgBox(Text, Title " Modded Msgbox", Options) {
		case "OK", "Yes": return 1
		case "No": return 2
		case "Cancel": return 3
	}
}

slowerevent(event, slow := 10) {
	ordelay := A_KeyDelay
	setkeydelay(slow, slow)
	SendEvent(event)
	setkeydelay(ordelay, ordelay)
}

mySelectInput(Type, Answers, Title := "mySelectInput", Prompt := "", options := "") {
	selectGUI := Gui(, Title), selectGUI.SetFont("S15"), answers.InsertAt(1, "")
	Prompt ? selectGUI.AddText(, Prompt) : ""
	selectOpt := selectGUI.Add(Type, "vres Multi r" ((length := Answers.length) <= 10 ? length : 10) " " options, answers)
	selectGUI.Show("AutoSize")
	WinWaitClose(selectGUI)
	if res := selectOpt.Text
		return res
	Exit
}

okinputbox(Prompt?, Title?, Options := "h100", Default?) {
	temp := InputBox(Prompt?, Title?, Options, Default?)
	if temp.result = "OK"
		return temp.Value
	Exit
}

FileOverwrite(Text,Filename){
	try FileDelete Filename
	FileAppend Text,Filename
}

ThousandsSep(x, s := ",") => RegExReplace(x, "\G\d+?(?=(\d{3})+(?:\D|$))", "$0" s)

