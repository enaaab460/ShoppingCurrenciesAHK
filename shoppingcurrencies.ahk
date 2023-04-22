SetWorkingDir A_WorkingDir
#Include "Lib\select toolbox.ahk"
DetectHiddenWindows 0

A_TrayMenu.add("Converter", (*) => convgui.show())
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
statusbar.OnEvent("Click", (obj, info) => (info = 1) ? settingsgui.Show("X1200") : "")

initiateini() {
    global
    SettingsIni := Yaml("shoppingcurrencies.yml")[1]["Settings"]
    baseCurrency := SettingsIni["Base"]
    intCurrency := SettingsIni["INT"]
    if (!FileExist("currency.yml") or !instr(FileGetTime("currency.yml", "M"), A_Year A_mon A_DD))
        Download "http://www.floatrates.com/daily/" baseCurrency ".json", "currency.yml"
    currencyjson := Yaml("currency.yml")
    usdrate := currencyjson[Strlower(intCurrency)]["inverseRate"]
    if custrate := SettingsIni["Alt_$"] > 0 {
        altfactor := custrate / usdrate
        usdrate := custrate
    } else {
        altfactor := 1 + SettingsIni["BankRate_%"] / 100
        usdrate := usdrate * altfactor
    }
    local lastcur := convgui["fromCur"].Text, currencylist := []
    for c in strsplit(SettingsIni["Regions"], ",")
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
            overhead := round(SettingsIni["Traveler_$"] * usdrate) + (direct * (SettingsIni["IntFees_%"] / 100))
        else if convgui["Overhead"].Text = "Shipping"
            overhead := round(direct * ((1 + SettingsIni["LocalFees_%"] / altfactor / 100) * ((SettingsIni["IntFees_%"]) / 100 + 1)) + SettingsIni["Shipping_$"] * usdrate) - direct
        else outformat := "{} {}", overhead := 0
        convgui["toVal"].Text := Format(outformat, toCur, ThousandsSep(round(direct)), ThousandsSep(round(direct + overhead)))
        if convgui["fromCur"].Text != baseCurrency and direct > SettingsIni["BankMax"] and !custrate
            failedmsg "Price exceeds maximum exchange allowed by bank, change to altrate"
    }
}

initiateini()

settingsgui := Gui()
settingsgui.SetFont("S18")
settingsnames := ["Currencies", "Base", "INT", "Regions", "Conversion", "BankRate_%", "BankMax", "Alt_$", "Overhead", "IntFees_%", "Traveler_$", "LocalFees_%", "Shipping_$"]
for editbox in settingsnames {
    settingsgui.AddText("xs y" A_Index * 40 - 36, editbox)
    settingsgui.Add(InStr("Currencies,Conversion,Overhead", Editbox) ? "Link" : "Edit", Format("x200 y{} h36 w150 v{}", A_Index * 40 - 36, editbox), SettingsIni[editbox])
}
settingsgui.AddButton("xs+120", "Save").OnEvent("Click", Saveset)
Saveset(*) {
    global SettingsIni
    for key, value in SettingsIni
        SettingsIni[key] := settingsgui[key].Text ? settingsgui[key].Text : 0
    fileoverwrite(Yaml(Map("Settings", SettingsIni), 2), "shoppingcurrencies.yml")
    calculateresult()
    settingsgui.Hide()
}

convgui.Show()
#HotIf WinActive("ahk_exe chrome.exe")
f7:: {
    convgui["fromVal"].Text := RegExReplace(copynow(, 0.2), "[^\d,.]"), calculateresult()
    convgui.Show()
    KeyWait(ThisHotkey, "t0.3")
}
f8:: {
    currentlink := UIA_Chrome("A").GetCurrentURL()
    for currency, stores in ymldb["Regions"]
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
            upperendS := "", upperendT := ""
        default: tCurrency := baseCurrency
            overheadmode := ymldb["Settings"]["Overhead"]["Mode"]
            overheadmode := overheadmode = "Ask" ? mySelectInput("DropDownList", ["Shipping", "Traveler", "Both", "None"], , "Select Conversion Mode, or keep empty to cancel") : overheadmode
            convMode := currencyjson[StrLower(MatchC)]["inverseRate"] * altfactor
            upperendS := Format('+ " - S " + Math.round(targetEGP* {1} *{2} + {3} * {1})', convMode, (1 + overheadS["LocalFees_%"] / altfactor / 100) * (overheadS["IntFees_%"] / 100 + 1), overheadS["Shipping_$"])
            upperendT := Format('+ " - T " + Math.round(targetEGP* {1} *1.{2} + {3} * {1})', convMode, overheadS["IntFees_%"], overheadS["Traveler_$"])
            switch overheadmode {
                case "Shipping": upperendT := ""
                case "Traveler": upperendS := ""
                case "Both": v := 1
                case "None": upperendS := "", upperendT := ""
            }
    }
    jscmd := Format('
    ( LTrim Join`s
        for (target of document.querySelectorAll("{1}")){
            targetEGP = target.innerText.replace(",","").match(/\d+/);if (targetEGP) targetEGP = targetEGP[0]; else continue;
            directconv = Math.round(targetEGP*{3});
            target.textContent += "\n({2} " + directconv {4} {5} + ")";
            if (directconv >= {6} && {7}) {window.alert("EGP " + directconv + " exceeds maximum exchange allowed by bank, change to altrate")}else{void(0);};
        }
    )', MatchQ, tCurrency, convMode, upperendS, upperendT, ymldb["Settings"]["Conversion"]["BankMax"], !(custrate or MatchC = baseCurrency) ? 1 : 0)
    UIA_Chrome("A").JSExecute(jscmd)
    ; savemsg jscmd
}
f9::
f10:: {
    chrome := UIA_Chrome("A")
    RegExMatch(chrome.GetCurrentURL(), "U)https?://(www\.)?(?<host>[\w.]+)\/", &urlregex)
    if ThisHotkey = "f10" {
        RegExMatch(copynow(, 0.2), "[\d.,]+", &regex)
        coo := regex[]
        firstjs := format("
        ( LTrim Join`s
        spanClassArray = new Set();
        for (span of document.querySelectorAll('{}')){
            if (span.innerText.search('{}') == -1 || (spanY = span.getBoundingClientRect().y) < 0 || spanY > 1040 || !(spanA = span.attributes[0])) continue;
            if (spanC = span.className) suffix = '.' + spanC.replaceAll(' ','.');
            else if (spanI = span.id) suffix = '#' + spanI;
            else suffix = '[' + spanA.name + '="' + spanA.value + '"]';
            spanClassArray.add(span.tagName.toLowerCase() + suffix);
        }
    )", "p,span,a,td,li,ul,strong", okinputbox("Write the price currently onscreen", , , coo))
    chrome.JSExecute(firstjs ';void(0);')
        spanArray := chrome.JSReturnThroughClipboard('Array.from(spanClassArray).join("|||")')
        spanarray := StrSplit(spanArray, "|||")
        loop spanarray.Length {
            jscmd := Format('for (span of document.querySelectorAll("{1}")) span.innerText = "SUCCESS";void(0);', span := spanarray[-A_Index])
            chrome.JSExecute(jscmd)
            if (res := ModMsg("Testing cssselector " span "`nIf target price changed to sucess, Choose `"Success`"`nIf page broke down, Choose `"Refresh`"`nOtherwise choose `"Next`"", , 4096, ["&Success", "&Refresh", "&Next"])) = 1 {
                targetspan := span
                break
            } else if res = 2 {
                chrome.Reload()
                chrome.WaitPageLoad()
            }
        }
    } else {
        WinActivate "ahk_exe chrome.exe"
        send "^C"
        ; sleep 1000
        chrome.WaitElement({ Name: "Toggle device toolbar", T: "Button" })
        ToolTip "Choose Price"
        KeyWait "LButton", "D"
        ToolTip
        sleep 50
        slowerevent("{AppsKey}{Sleep 1}{c 2}{Right}{c 2}{Enter}", 100)
        regex := RegExMatchAll(A_Clipboard, "\S*[.#]\S+")
        try targetspan := regex[-1][]
    }
    if isset(targetspan) {
        if MsgBox("Do you add this store before?", , 4) = "No"
            res := savemsg(A_Tab A_Tab strreplace(urlregex["host"], "www.") ': ' targetspan, "Press yes to copy, then paste the following to the correct region in the yml file`n")
        else res := savemsg(', ' targetspan, "Press yes to copy, then paste the following to the correct store in the yml file`n")
        res = 0 ? "" :settings()
    } else
        MsgBox("Failed to find cssselector")
}