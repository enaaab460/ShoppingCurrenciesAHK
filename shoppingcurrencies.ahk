SetWorkingDir A_WorkingDir
#Include "<select toolbox>"
DetectHiddenWindows 0

utcdif := DateDiff(A_Now, A_NowUTC, "H")
; moddate := dateadd(FileGetTime(A_ScriptName, "M"),-utcdif,"h")
moddate := 20230520142714
moddate := dateadd(moddate, -utcdif, "h")
repoupdate := 0

A_TrayMenu.add("Converter", (*) => convgui.show())
A_TrayMenu.add("Settings", (*) => settingsgui.show())
A_TrayMenu.default := "Converter"
A_TrayMenu.ClickCount := 1

convgui := Gui()
convgui.SetFont("s18")
convgui.AddText(, "From Currency")
convgui.AddDropDownList("vfromCur ys x250 w120").OnEvent("Change", calculateresult)
convgui.AddEdit("vfromVal xm w350").OnEvent("Change", calculateresult)
convgui.AddText("xm voverheadtext section", "Overhead Mode")
convgui.AddDropDownList("vOverhead yp x250 w120 choose1", ["Convert", "Shipping", "Traveler", "T+T"]).OnEvent("Change", calculateresult)
; convgui.AddEdit("vtoVal ReadOnly r1 xm w350")
convgui.AddComboBox("vtoVal xm w350")
(statusbar := convgui.AddStatusBar('vStatus')).SetParts(20, 20, 240)
statusbar.SetIcon(A_WinDir "\System32\" "dsuiext.dll", 36)
statusbar.SetIcon(A_WinDir "\System32\" "comres.dll", 7, 3)
statusbar.SetIcon("lib\youtube.png", , 2)
statusbar.OnEvent("Click", (obj, info) => info = 1 ? (settingsgui.Show(), convgui.Hide()) : info = 2 ? run("https://youtu.be/qP5hBoRbKWc") : info = 3 ? editcurgui.Show() : "")

InitiateYml() {
    global
    SettingsYml := Yaml("settings.yml")[1]
    baseCurrency := SettingsYml["Base"], intCurrency := SettingsYml["INT"]
    if !FileExist("currency.json") or !instr(FileGetTime("currency.json", "M"), A_Year A_mon A_DD) or InStr(FileRead("currency.json"), baseCurrency)
        try Download("http://www.floatrates.com/daily/" baseCurrency ".json", "currency.json"), repoupdate := regexreplace(Yaml(httprequest('https://api.github.com/repos/enaaab460/ShoppingCurrenciesAHK/branches/master2'))['commit']['commit']['author']['date'], "\D")
        catch
            MsgBox("Failed to connect to online sources`nExiting Program", , "Icon!"), ExitApp()
    curfile := FileRead('currency.json')
    for i, match in RegExMatchAll(curfile, '([\d.]+)e-(\d)')
        curfile := StrReplace(curfile, match[0], Format('{:.13f}', (match[1]) / (10 ** match[2])))
    currencyjson := Yaml(curfile)
    baseCurrency = intCurrency ? currencyjson[Strlower(intCurrency)] := Map("inverseRate", 1) : ""
    bankusd := round(currencyjson[Strlower(intCurrency)]["inverseRate"], 2)
    altusd := SettingsYml["Alt_$"] > 0 ? SettingsYml["Alt_$"] : 0
    altfactor := altusd ? altusd / bankusd : 1
    bankmax := SettingsYml["BankMax"]
    if InStr(bankmax, "$")
        bankmax := StrReplace(bankmax, "$") * bankusd
    local lastcur := convgui["fromCur"].Text
    currencylist := strsplit(SettingsYml["Regions"], ",")
    convgui["fromCur"].Delete(), convgui["fromCur"].Add(currencylist)
    try convgui["fromCur"].Choose(lastcur != "" ? lastcur : intCurrency)
    catch
        convgui["fromCur"].Choose(intCurrency)
}

calculateresult(*) {
    if convgui["fromCur"].Text = baseCurrency {
        toCur := intCurrency
        convrate := 1 / (altusd > 0 ? altusd : bankusd)
        ; outformat := "{} {}"
        ; convgui["Overhead"].Enabled := 0, convgui["Overhead"].Text := "Convert"
        statusbar.SetText(altusd ? "Alt Rate" : "", 4)
        maxout := bankmax
    } else {
        toCur := baseCurrency
        convrate := currencyjson[StrLower(convgui["fromCur"].Text)]["inverseRate"] * altfactor
        ; outformat := "{} {}-{}"
        ; convgui["Overhead"].Enabled := 1
        maxout := ((bankmax / convrate) - (SettingsYml["ShipNow"] and convgui["Overhead"].Text = "Shipping" ? SettingsYml["Shipping"] : 0)) * ((100 + abs(SettingsYml["IntFees_%"])) / 100) ** (SettingsYml["IntFees_%"] > 0 ? -1 : 1)
        statusbar.SetText(altusd ? "Alt Rate" : floor(maxout), 4)
    }
    outformat := "{} {}-{}"
    statusbar.SetText("1 " convgui["fromCur"].Text " = " round(convrate, 2) " " toCur, 3)
    ; statusbar.SetText("1 " convgui["fromCur"].Text " = " format('{:.2g}',convrate) " " toCur, 3)
    convgui["toVal"].Delete()
    if convgui["fromVal"].Text ~= "[^\d.,]"
        convgui["toVal"].add(["Letters,commas and spaces not allowed"])
    else if !convgui["fromVal"].Text
        convgui["toVal"].value := ""
    else {
        ; out := regexreplace(convgui["fromVal"].Text, "\.\d*|,")
        out := strreplace(convgui["fromVal"].Text, ",")
        if toCur = baseCurrency {
            outaIntFees := out * ((1 + (abs(SettingsYml["IntFees_%"]) / 100)) ** (SettingsYml["IntFees_%"] >= 0 ? 1 : -1))
            convtoutaIntFees := outaIntFees * convrate
            bankcomm := convtoutaIntFees * (altfactor > 1 ? 0 : SettingsYml["BankRate_%"] / 100)
            switch convgui["Overhead"].Text {
                case "Convert": transportcomm := 0
                case "Traveler", "T+T": transportcomm := instr(SettingsYml["Traveler_$"], "a") ? StrReplace(SettingsYml["Traveler_$"], "a") * altusd : SettingsYml["Traveler_$"] * bankusd
                case "Shipping": transportcomm := SettingsYml["Shipping"] * (InStr(SettingsYml["Shipping"], "$") ? strreplace(bankusd, "$") : convrate)
                    bankcomm += transportcomm * SettingsYml["BankRate_%"] / 100
            }
            localfeecent := convgui["Overhead"].Text = "Traveler" ? 0 : outaIntFees * (convrate / altfactor) * SettingsYml["LocalFees_%"] / 100
            total := convtoutaIntFees + (overhead := bankcomm + (transportcomm ?? 0) + localfeecent + SettingsYml["LocalFees"])
            direct := convtoutaIntFees + (SettingsYml["ShipNow"] and convgui["Overhead"].Text = "Shipping" ? transportcomm : 0)
            convgui["toVal"].add([Format(outformat, toCur, ThousandsSep(round(direct)), ThousandsSep(round(total))), "Price after tax: " ThousandsSep(round(convtoutaIntFees, 2))])
            transportcomm ? convgui["toVal"].add(["Trasport Cost: " ThousandsSep(round(transportcomm, 2))]) : ''
            (bankcomm > 0) ? convgui["toVal"].add(["Bank Commision: " ThousandsSep(round(bankcomm, 2))]) : ''
            (localfeecent) or SettingsYml["LocalFees"] > 0 ? convgui["toVal"].add(["Local Fees: " ThousandsSep(round(localfeecent + SettingsYml["LocalFees"], 2))]) : ""
            ; if outaIntFees > maxout and !altusd
            if direct > bankmax and !altusd
                MsgBox ("Price exceeds maximum exchange allowed by bank , change to altrate"), , 48
        } else {
            ; convgui["toVal"].add([Format(outformat, toCur, ThousandsSep(round(out * convrate))), Format("Before Fees: {} {}", toCur, beforefees := round((out - SettingsYml['LocalFees']) / (1 + SettingsYml['LocalFees_%'] / 100) / (1 + SettingsYml['IntFees_%'] / 100) * convrate / (1 + SettingsYml['BankRate_%'] / 100)))])
            ; toValList := []
            ; convgui["toVal"].add([Format(outformat, toCur, ThousandsSep(round(out * convrate)))])
            ; switch convgui["Overhead"].Text {
            ;     case "Convert": convgui["toVal"].add([Format("Before Fees: {} {}", toCur, ThousandsSep(round((out - SettingsYml['LocalFees']) / (1 + SettingsYml['LocalFees_%'] / 100) / (1 + SettingsYml['IntFees_%'] / 100) * convrate / (1 + SettingsYml['BankRate_%'] / 100))))])
            ;     case "Traveler": convgui["toVal"].add([Format("Before Traveler: {} {}", toCur, ThousandsSep(Round(((out - SettingsYml['LocalFees']) * convrate - SettingsYml['Traveler_$']) / (1 + SettingsYml['IntFees_%'] / 100) / (1 + SettingsYml['BankRate_%'] / 100), 2)))])
            ;     case "Shipping": convgui["toVal"].add([Format("Before Shipping: {} {}", toCur, ThousandsSep(Round(((out - SettingsYml['LocalFees']) * convrate - ((1 + SettingsYml['BankRate_%'] / 100) * SettingsYml['Shipping'])) / (((1+SettingsYml['BankRate_%']/100) + (SettingsYml['LocalFees_%']/100)) * (1+SettingsYml['IntFees_%']/100)), 2)))])
            ; }
            switch convgui["Overhead"].Text {
                case "Convert": before := (out - SettingsYml['LocalFees']) / (1 + SettingsYml['LocalFees_%'] / 100) / (1 + SettingsYml['IntFees_%'] / 100) * convrate / (1 + SettingsYml['BankRate_%'] / 100)
                case "Traveler": before := ((out - SettingsYml['LocalFees']) * convrate - SettingsYml['Traveler_$']) / (1 + SettingsYml['IntFees_%'] / 100) / (1 + SettingsYml['BankRate_%'] / 100)
                case "T+T": before := 0
                case "Shipping": before := ((out - SettingsYml['LocalFees']) * convrate - ((1 + SettingsYml['BankRate_%'] / 100) * SettingsYml['Shipping'])) / (((1 + SettingsYml['BankRate_%'] / 100) + (SettingsYml['LocalFees_%'] / 100)) * (1 + SettingsYml['IntFees_%'] / 100))
            }
            convgui["toVal"].add([Format(outformat, toCur, ThousandsSep(round(out * convrate)), ThousandsSep(round(before)))])
        }
        convgui["toVal"].choose(1)
    }
}

InitiateYml()
calculateresult()
convgui.Show()

settingsgui := Gui()
settingsgui.SetFont("S18")
settingsnames := ["Currencies", "Base", "INT", "Regions", "Conversion", "BankRate_%", "BankMax", "Alt_$", "Overhead", "IntFees_%", "Traveler_$", "Shipping", "LocalFees_%", "LocalFees", "F8Mode"]
for editbox in settingsnames {
    settingsgui.AddText("xs y" (A_Index - 1) * 38, editbox), inputformat := Format("x200 yp h36 w150 v{}", editbox)
    ; editbox = "BankMax" ? settingsgui.AddPicture("x+20 w20 h20 icon95", A_WinDir "\System32\imageres.dll").OnEvent("Click", (*) => run(SettingsYml["BankInfo"])) : ""
    editbox = "Shipping" ? settingsgui.AddCheckbox("vShipNow y" (A_Index - 1) * 38 " x+5 " (SettingsYml["ShipNow"] ? "Checked" : ""), "Now") : ""
    if instr(editbox, "F8Mode")
        settingsgui.Add("DropDownList", inputformat " r5", ["Ask", "Convert", "Shipping", "Traveler", "T+T"]).Text := SettingsYml[editbox]
    else settingsgui.Add(InStr("Currencies,Conversion,Overhead", Editbox) ? "Link" : "Edit", inputformat, SettingsYml[editbox])
}
settingsgui.AddButton("y+10 xm+120", "Save").OnEvent("Click", Saveset)
settingsgui.OnEvent("Close", (*) => convgui.Show())
settingsstatus := settingsgui.AddStatusBar()
settingsstatus.SetParts(20, 20)
settingsstatus.OnEvent("Click", (obj, info) => info = 2 ? Run("notepad.exe stores.yml") : run("https://github.com/enaaab460/ShoppingCurrenciesAHK"))
settingsstatus.SetIcon("lib\github.png", , 1), settingsstatus.SetIcon("lib\store.png", , 2)
settingsstatus.SetText(repoupdate > moddate ? "Update available" : "Version:" FormatTime(moddate, 'shortdate'), 3)

Saveset(*) {
    global SettingsYml
    for key, value in SettingsYml
        SettingsYml[key] := settingsgui[key].Text ? settingsgui[key].Text : 0
    SettingsYml["ShipNow"] := settingsgui["ShipNow"].Value
    FileOverwrite(Yaml(SettingsYml, 2), "settings.yml")
    InitiateYml()
    calculateresult()
    convgui.Show()
    settingsgui.Hide()
}

editcurgui := Gui()
editcurgui.SetFont("S14")
editcurgui.AddRadio("vCur Checked", baseCurrency)
editcurgui.AddRadio("x+50", convgui["fromCur"].Text)
editcurgui.AddEdit("y+5 xm w200 vCurValue")
editcurgui.AddButton("y+5 xm+60", "Save").OnEvent("Click", editcur)
editcur(*) {
    input := editcurgui.Submit(1)
    (input.cur = 1) ? (input.curvalue := 1 / input.curvalue) : ""
    currencyjson[strlower(convgui["fromCur"].Text)]["inverseRate"] := input.curvalue
    calculateresult()
}

#HotIf WinActive("ahk_exe chrome.exe")

f7:: {
    static currencysymbols := Yaml("Common-Currency.json")
    A_Clipboard := cn := StrReplace(copynow(, 0.2), "`r`n")
    if RegExMatch(cn, "[^\d., ]+", &regex)
        for cur in currencylist
            if regex[0] = "$"
                convgui["fromCur"].choose("USD")
            else if regex[0] = currencysymbols[cur]["symbol"] or regex[0] = currencysymbols[cur]["symbol_native"]
                match := convgui["fromCur"].choose(cur)
    cn := RegExReplace(cn, ",(\d{2})\D+", ".$1")
    convgui["fromVal"].Text := RegExReplace(cn, "[^\d,.]"), calculateresult()
    convgui.Show()
    KeyWait(ThisHotkey, "t0.3")
}

f8::
chromeprice(*) {
    WinActivate "ahk_exe chrome.exe"
    chrome := UIA_Chrome("ahk_exe chrome.exe")
    chrome.ElementExist({ Name: "Toggle device toolbar", T: "Button" }) ? Send("F12{Sleep 300}") : ""
    currentlink := chrome.GetCurrentURL()
    for currency, stores in Yaml("stores.yml")[1]
        if stores
            for store, query in stores {
                RegExMatch(currentlink, store, &regex)
                if regex {
                    matchQ := query
                    MatchC := currency
                    break 2
                }
            }
    failedmsg("Store not supported", !isset(MatchC))
    switch MatchC {
        case baseCurrency: toCur := intCurrency
            convrate := 1 / (altusd > 0 ? altusd : bankusd)
            overheadmode := "Convert"
            minijs := Format('target.textContent += "\n({} " + Math.round(productPrice * {}) + ")";', toCur, convrate)
        default: toCur := baseCurrency
            convrate := currencyjson[StrLower(MatchC)]["inverseRate"] * altfactor
            overheadmode := SettingsYml["F8Mode"] = "Ask" ? mySelectInput("DropDownList", ["Convert", "Shipping", "Traveler", "TT"], , "Select Conversion Mode, or keep empty to cancel") : SettingsYml["F8Mode"]
            switch overheadmode {
                case "Convert": mmath := ""
                case "Traveler", "TT": mmath := Format('transportcomm = {};', transportcomm := instr(SettingsYml["Traveler_$"], "a") ? StrReplace(SettingsYml["Traveler_$"], "a") * altusd : SettingsYml["Traveler_$"] * bankusd)
                case "Shipping": mmath := Format('transportcomm = {};`nbankcomm += {};', transportcomm := SettingsYml["Shipping"] * (InStr(SettingsYml["Shipping"], "$") ? strreplace(bankusd, "$") : convrate), transportcomm * SettingsYml["BankRate_%"] / 100)
            }
            minijs := Format('
            (
                intFees = productPrice * {};
                outaIntFees = intFees + productPrice;
                convtoutaIntFees = outaIntFees * {};
                bankcomm = convtoutaIntFees * {};
                {}
                localfeecent = {};
                total = convtoutaIntFees + bankcomm + {} + localfeecent;
                target.textContent += "\n({} " + Math.round(convtoutaIntFees) + "-" + "{}" + Math.round(total) + ")";
                if (convtoutaIntFees >= {} && {}) {window.alert("EGP " + Math.round(convtoutaIntFees) + " exceeds maximum exchange allowed by bank, change to altrate")}else{void(0);};
            )', SettingsYml["IntFees_%"] >= 0 ? SettingsYml["IntFees_%"] / 100 : -(SettingsYml["IntFees_%"] / (100 + SettingsYml["IntFees_%"])), convrate, (altfactor > 1 ? 0 : SettingsYml["BankRate_%"] / 100), mmath, overheadmode = "Traveler" ? 0 : "outaIntFees *" convrate * SettingsYml["LocalFees_%"] / 100, (transportcomm ?? 0) + SettingsYml["LocalFees"], toCur, SubStr(overheadmode, 1, 2), bankmax - (SettingsYml["ShipNow"] and overheadmode = "Shipping" ? transportcomm : 0), !(altusd) ? 1 : 0)
    }
    jscmd := Format('
    ( ;LTrim Join`s
        for (target of document.querySelectorAll("{}")){
            productPrice = target.innerText.replace(",","").replaceAll("\n","").match(/[\d.]+/);if (productPrice) productPrice = Number(productPrice[0]); else continue;
            {}
        }void(0);
    )', MatchQ, minijs)
    WinActivate "ahk_exe chrome.exe"
    chrome.JSExecute(jscmd)
    ; savemsg jscmd
}

f9:: {
    WinActivate "ahk_exe chrome.exe"
    chrome := UIA_Chrome("A")
    RegExMatch(chrome.GetCurrentURL(), "U)https?:\/\/(?:www\.)?(?<host>[\w.]+)\/", &urlregex)
    send "^C"
    sleep 400
    chrome.WaitElement({ Name: "Toggle device toolbar", T: "Button" })
    ToolTip "Choose Price"
    KeyWait "LButton", "D"
    ToolTip
    sleep 50
    A_Clipboard := ""
    slowerevent("{AppsKey}{Sleep 1}{c 2}{Right}{c 2}{Enter}", 100)
    ClipWait(1)
    regex := RegExMatchAll(A_Clipboard, "\S*[.#]\S+")
    try targetspan := regex[-1][]
    failedmsg "Failed to find cssselector", !IsSet(targetspan)
    ToolTip targetspan
    cur := StrUpper(mySelectInput("ComboBox", currencylist, , "Currency of store"))
    store := okinputbox("Confirm that the url is correct and specific to region", , , strreplace(urlregex["host"], "www."))
    regionyml := Yaml("stores.yml")[1]
    regionyml.has(cur) ? "" : regionyml[cur] := Map()
    if regionyml[cur].has(store) and InStr(regionyml[cur][store], targetspan)
        MsgBox(Format("{} is already a part of {}'s css selectors`nDo you want to replace the whole css?", targetspan, store), , 4) = "Yes" ? regionyml[cur][store] := targetspan : ""
    else regionyml[cur][store] := targetspan
    FileOverwrite(Yaml(regionyml, 3), "stores.yml")
    ToolTip
    WinActivate "ahk_exe chrome.exe"
    send "{F12}{Sleep 200}"
    chromeprice()
}