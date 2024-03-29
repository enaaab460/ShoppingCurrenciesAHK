SetWorkingDir A_WorkingDir
#Include "Lib\select toolbox.ahk"
DetectHiddenWindows 0

updatedate := '24/4/23'

A_TrayMenu.add("Converter", (*) => convgui.show())
A_TrayMenu.default := "Converter"
A_TrayMenu.ClickCount := 1

convgui := Gui()
convgui.SetFont("s18")
convgui.AddText(, "From Currency")
convgui.AddDropDownList("vfromCur ys x250 w120").OnEvent("Change", calculateresult)
convgui.AddEdit("vfromVal xm w350").OnEvent("Change", calculateresult)
convgui.AddText("xm voverheadtext section", "Overhead Mode")
convgui.AddDropDownList("vOverhead yp x250 w120 choose1", ["Convert", "Shipping", "Traveler", "T+T"]).OnEvent("Change", calculateresult)
convgui.AddEdit("vtoVal ReadOnly r1 xm w350")
(statusbar := convgui.AddStatusBar('vStatus')).SetParts(20, 20, 240)
statusbar.SetIcon(A_WinDir "\System32\" "dsuiext.dll", 36)
statusbar.SetIcon("lib\youtube.png", , 2)
statusbar.OnEvent("Click", (obj, info) => info = 1 ? (settingsgui.Show(), convgui.Hide()) : info = 2 ? run("https://youtu.be/qP5hBoRbKWc") : "")

InitiateYml() {
    global
    SettingsYml := Yaml("settings.yml")[1]
    baseCurrency := SettingsYml["Base"]
    intCurrency := SettingsYml["INT"]
    if (!FileExist("currency.json") or !instr(FileGetTime("currency.json", "M"), A_Year A_mon A_DD))
        Download "http://www.floatrates.com/daily/" baseCurrency ".json", "currency.json"
    currencyjson := Yaml("currency.json")
    usdrate := currencyjson[Strlower(intCurrency)]["inverseRate"]
    if (custrate := SettingsYml["Alt_$"]) > 0 {
        altfactor := custrate / usdrate
        usdrate := custrate
        statusbar.SetText("Alt Rate", 4)
    } else {
        altfactor := 1 + SettingsYml["BankRate_%"] / 100
        usdrate := usdrate * altfactor
    }
    bankmax := SettingsYml["BankMax"]
    if InStr(bankmax, "$")
        bankmax := StrReplace(bankmax, "$") * usdrate / altfactor
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
        convrate := 1 / (usdrate / (custrate > 0 ? 1 : altfactor))
        outformat := "{} {}"
        convgui["Overhead"].Enabled := 0
        statusbar.SetText(custrate > 0 ? "Alt Rate" : bankmax, 4)
    } else {
        toCur := baseCurrency
        convrate := currencyjson[StrLower(convgui["fromCur"].Text)]["inverseRate"] * altfactor
        outformat := "{} {}-{}"
        convgui["Overhead"].Enabled := 1
        statusbar.SetText(custrate ? "Alt Rate" : round(bankmax / convrate * altfactor), 4)
    }
    statusbar.SetText("1 " convgui["fromCur"].Text " = " round(convrate, 2) " " toCur, 3)
    if convgui["fromVal"].Text ~= "[a-zA-Z]"
        convgui["toVal"].Text := "Letters,commas and spaces not allowed"
    else if !convgui["fromVal"].Text
        convgui["toVal"].Text := ""
    else {
        out := regexreplace(convgui["fromVal"].Text, "\.\d*|,")
        direct := round(out * convrate, 2)
        if convgui["Overhead"].Text = "Traveler"
            overhead := round(SettingsYml["Traveler_$"] * usdrate) + (direct * (SettingsYml["IntFees_%"] / 100)) + SettingsYml["LocalFees"]
        else if convgui["Overhead"].Text = "Shipping"
            overhead := round(direct * ((1 + SettingsYml["LocalFees_%"] / altfactor / 100) * ((SettingsYml["IntFees_%"]) / 100 + 1)) + SettingsYml["Shipping_$"] * usdrate) - direct + SettingsYml["LocalFees"]
        else if convgui["Overhead"].Text = "T+T"
            overhead := round(direct * ((1 + SettingsYml["LocalFees_%"] / altfactor / 100) * ((SettingsYml["IntFees_%"]) / 100 + 1)) + SettingsYml["Traveler_$"] * usdrate) - direct + SettingsYml["LocalFees"]
        else outformat := "{} {}", overhead := 0
        convgui["toVal"].Text := Format(outformat, toCur, ThousandsSep(round(direct)), ThousandsSep(round(direct + overhead)))
        if toCur = baseCurrency and direct > bankmax * altfactor and !custrate
            MsgBox "Price exceeds maximum exchange allowed by bank, change to altrate", , 48
    }
}

InitiateYml()
calculateresult()
convgui.Show()

settingsgui := Gui()
settingsgui.SetFont("S18")
settingsnames := ["Currencies", "Base", "INT", "Regions", "Conversion", "BankRate_%", "BankMax", "Alt_$", "Overhead", "IntFees_%", "Traveler_$", "Shipping_$", "LocalFees_%", "LocalFees", "F8Mode"]
for editbox in settingsnames {
    settingsgui.AddText("xs y" A_Index * 40 - 36, editbox), inputformat := Format("x200 yp h36 w150 v{}", editbox)
    ; editbox = "BankMax" ? settingsgui.AddPicture("x+20 w20 h20 icon95", A_WinDir "\System32\imageres.dll").OnEvent("Click", (*) => run(SettingsYml["BankInfo"])) : ""
    if instr(editbox, "F8Mode")
        settingsgui.Add("DropDownList", inputformat " r5", ["Ask", "Convert", "Shipping", "Traveler", "T+T"]).Text := SettingsYml[editbox]
    else settingsgui.Add(InStr("Currencies,Conversion,Overhead", Editbox) ? "Link" : "Edit", inputformat, SettingsYml[editbox])
}
settingsgui.AddButton("xs+120", "Save").OnEvent("Click", Saveset)
settingsgui.OnEvent("Close", (*) => convgui.Show())
settingsstatus := settingsgui.AddStatusBar()
settingsstatus.SetParts(20,20)
settingsstatus.OnEvent("Click", (obj,info) =>info = 2 ? Run("notepad.exe stores.yml") : run("https://github.com/enaaab460/ShoppingCurrenciesAHK"))
settingsstatus.SetIcon("lib\github.png", , 1)
settingsstatus.SetIcon("lib\store.png", , 2)
settingsstatus.SetText("Updated on: " updatedate,3)
Saveset(*) {
    global SettingsYml
    for key, value in SettingsYml
        SettingsYml[key] := settingsgui[key].Text ? settingsgui[key].Text : 0
    FileOverwrite(Yaml(SettingsYml, 2), "settings.yml")
    InitiateYml()
    calculateresult()
    convgui.Show()
    settingsgui.Hide()
}

#HotIf WinActive("ahk_exe chrome.exe")
f7:: {
    convgui["fromVal"].Text := RegExReplace(copynow(, 0.2), "[^\d,.]"), calculateresult()
    convgui.Show()
    KeyWait(ThisHotkey, "t0.3")
}

f8::
chromeprice(*) {
    WinActivate "ahk_exe chrome.exe"
    currentlink := UIA_Chrome("ahk_exe chrome.exe").GetCurrentURL()
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
    isset(MatchC) ? "" : failedmsg("Store not supported")
    switch MatchC {
        case baseCurrency: tCurrency := intCurrency
            convMode := 1 / (usdrate / (custrate > 0 ? 1 : altfactor))
            upperend := ""
        default: tCurrency := baseCurrency
            overheadmode := SettingsYml["F8Mode"] = "Ask" ? mySelectInput("DropDownList", ["Convert", "Shipping", "Traveler", "TT"], , "Select Conversion Mode, or keep empty to cancel") : SettingsYml["F8Mode"]
            convMode := currencyjson[StrLower(MatchC)]["inverseRate"] * altfactor
            switch overheadmode {
                case "Convert": upperend := ""
                case "Shipping": upperend := Format('+ " - S " + Math.round(targetEGP* {1} *{2} + {3} * {1})', convMode, (1 + SettingsYml["LocalFees_%"] / altfactor / 100) * (SettingsYml["IntFees_%"] / 100 + 1), SettingsYml["Shipping_$"])
                case "Traveler": upperend := Format('+ " - T " + Math.round(targetEGP* {1} *1.{2} + {3} * {1})', convMode, SettingsYml["IntFees_%"], SettingsYml["Traveler_$"])
                case "T+T": upperend := Format('+ " - TT " + Math.round(targetEGP* {1} *{2} + {3} * {1})', convMode, (1 + SettingsYml["LocalFees_%"] / altfactor / 100) * (SettingsYml["IntFees_%"] / 100 + 1), SettingsYml["Traveler_$"])
            }
    }
    jscmd := Format('
    ( LTrim Join`s
        for (target of document.querySelectorAll("{}")){
            targetEGP = target.innerText.replace(",","").match(/\d+/);if (targetEGP) targetEGP = targetEGP[0]; else continue;
            directconv = Math.round(targetEGP*{});
            target.textContent += "\n({} " + directconv {} + ")";
            if (directconv >= {} && {}) {window.alert("EGP " + directconv + " exceeds maximum exchange allowed by bank, change to altrate")}else{void(0);};
        }
    )', MatchQ, convMode, tCurrency, upperend, bankmax, !(custrate or MatchC = baseCurrency) ? 1 : 0)
    UIA_Chrome("A").JSExecute(jscmd)
    ; savemsg jscmd
}

f9:: {
    WinActivate "ahk_exe chrome.exe"
    chrome := UIA_Chrome("ahk_exe chrome.exe")
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
    cur := StrUpper(mySelectInput("ComboBox", currencylist, , "Currency of store"))
    store := okinputbox("Confirm that the url is correct and specific to region", , , strreplace(urlregex["host"], "www."))
    regionyml := Yaml("stores.yml")[1]
    regionyml.has(cur) ? "" : regionyml[cur] := Map()
    if regionyml[cur].has(store) and InStr(regionyml[cur][store], targetspan)
        MsgBox(Format("{} is already a part of {}'s css selectors`nDo you want to replace the whole css?", targetspan, store), , 4) = "Yes" ? regionyml[cur][store] := targetspan : ""
    else regionyml[cur][store] := targetspan
    FileOverwrite(Yaml(regionyml, 3), "stores.yml")
    WinActivate "ahk_exe chrome.exe"
    send "{F12}{Sleep 200}"
    chromeprice()
}