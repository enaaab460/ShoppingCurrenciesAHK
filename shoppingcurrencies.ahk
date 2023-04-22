SetWorkingDir A_WorkingDir
#Include "Lib\select toolbox.ahk"

A_TrayMenu.add("Converter", (*) => convgui.show())
A_TrayMenu.add("Settings", settings)
A_TrayMenu.default := "Converter"
A_TrayMenu.ClickCount := 1

convgui := Gui()
convgui.SetFont("s18")
convgui.AddText(, "From Currency")
convgui.AddDropDownList("vfromCur ys x250 w120").OnEvent("Change", calculateresult)
convgui.AddEdit("vfromVal xm w350").OnEvent("Change", calculateresult)
convgui.AddText("xm voverheadtext section", "Overhead Mode")
convgui.AddDropDownList("vOverhead yp x250 w120 choose1", ["None", "Shipping", "Traveler"]).OnEvent("Change", calculateresult)
convgui.AddEdit("vtoVal ReadOnly r2 xm w350")
(statusbar := convgui.AddStatusBar('vStatus')).SetParts(30)
statusbar.SetIcon(A_WinDir "\System32\" "dsuiext.dll", 36)
statusbar.OnEvent("Click", (obj, info) => (info = 1) ? settings() : "")

initiateini() {
    global
    ini := "shoppingcurrencies.ini"
    overheadS := Map()
    for itm in ["Traveler_$", "LocalFees_%", "IntFees_%", "Shipping_$"]
        overheadS[itm] := IniRead(ini, "Overhead", itm)
    baseCurrency := IniRead(ini, "Currencies", "Base")
    intCurrency := IniRead(ini, "Currencies", "INT")
    if (!FileExist("currency.yml") or !instr(FileGetTime("currency.yml", "C"), A_Year A_mon A_DD))
        Download "http://www.floatrates.com/daily/" baseCurrency ".json", "currency.yml"
    currencyjson := Yaml("currency.yml")
    usdrate := currencyjson[Strlower(intCurrency)]["inverseRate"]
    if custrate := IniRead(ini, "Conversion", "Alt_$") > 0 {
        altfactor := custrate / usdrate
        usdrate := custrate
    } else {
        altfactor := 1 + IniRead(ini, "Conversion", "BankRate_%") / 100
        usdrate := usdrate * altfactor
    }
    local lastcur := convgui["fromCur"].Text, currencylist := []
    for c in strsplit(IniRead(ini, "Currencies", "Regions"), ",")
        currencylist.Push(c)
    convgui["fromCur"].Delete(), convgui["fromCur"].Add(currencylist)
    try convgui["fromCur"].Choose(lastcur != "" ? lastcur : intCurrency)
    catch
        convgui["fromCur"].Choose(intCurrency)
    statusbar.SetText(" 1 " intCurrency " = " round(usdrate, 2) " " baseCurrency, 2)
}

calculateresult(*) {
    if convgui["fromCur"].Text = baseCurrency {
        toCur := intCurrency
        convrate := 1 / usdrate
        outformat := "{} {}"
        convgui["Overhead"].Enabled := 0
    } else {
        toCur := baseCurrency
        convrate := currencyjson[StrLower(convgui["fromCur"].Text)]["inverseRate"] * altfactor
        outformat := "{} {}-{}"
        convgui["Overhead"].Enabled := 1
    }
    if convgui["fromVal"].Text ~= "[a-zA-Z]"
        convgui["toVal"].Text := "Letters,commas and spaces not allowed"
    else if !convgui["fromVal"].Text
        convgui["toVal"].Text := ""
    else {
        out := regexreplace(convgui["fromVal"].Text, "\.\d*|,")
        direct := round(out * convrate, 2)
        if convgui["Overhead"].Text = "Traveler"
            overhead := round(overheadS["Traveler_$"] * usdrate) + (direct * (overheadS["IntFees_%"] / 100))
        else if convgui["Overhead"].Text = "Shipping"
            overhead := round(direct * ((1 + overheadS["LocalFees_%"] / altfactor / 100) * ((overheadS["IntFees_%"]) / 100 + 1)) + overheadS["Shipping_$"] * usdrate) - direct
        else outformat := "{} {}", overhead := 0
        convgui["toVal"].Text := Format(outformat, toCur, ThousandsSep(round(direct)), ThousandsSep(round(direct + overhead)))
        if convgui["fromCur"].Text != baseCurrency and direct > IniRead(ini, "Conversion", "BankMax") and !custrate
            failedmsg "Price exceeds maximum exchange allowed by bank, change to altrate"
    }
}



settingsgui := Gui()
settingsgui.SetFont("S18")
for label in ["Base Currency", "Int Currency", "Regions", "Bank Rate", "Bank Max", "Alt $","Int Fees %","Traveler $","Local Fees %","Shipping $"]
    settingsgui.AddText("xs", label)
for editbox in ["Base","Int","Regions"]
    settingsgui.AddEdit("x200 y+10 v" editbox)

settings(*) {
    settingsgui.Show("h400 w400")
    WinWaitClose settingsgui
    initiateini()
    calculateresult()
}

initiateini()
convgui.Show()
#HotIf WinActive("ahk_exe chrome.exe")
f7:: {
    convgui["fromVal"].Text := RegExReplace(copynow(, 0.2), "[^\d,.]"), calculateresult()
    convgui.Show()
    KeyWait(ThisHotkey, "t0.3")
}