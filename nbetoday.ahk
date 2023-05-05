#Include "lib\select toolbox.ahk"
run "chrome.exe https://www.nbe.com.eg/NBE/E/#/EN/ExchangeRatesAndCurrencyConverter"
currencies := OnlineJson("http://www.floatrates.com/daily/EGP.json")
WinWaitActive("ahk_exe chrome.exe")
chrome := UIA_chrome("A")
chrome.WaitPageLoad("National Bank of Egypt - Exchange Rates And Currency Converter",,,2)
nbetable := strsplit(chrome.JSReturnThroughClipboard('Array.from(document.querySelectorAll(".divExchangeRateCurrencyConverterExchangeRate  td#Banknote tr:nth-child(2) td:nth-child(2)")).map(x=>x.innerText).join(",")'), ",")
for code, index in Map("usd", 1, "eur", 2, "gbp", 3, "jpy", 9, "sar", 12, "aed", 13)
    currencies[code] := Map("inverseRate", nbetable[index])
currencies["jpy"]["inverseRate"] := round(currencies["jpy"]["inverseRate"] / 100, 4)
currencies["cny"] := map("inverseRate", chrome.JSReturnThroughClipboard('document.querySelectorAll(".page.divCurrencyConverterContainer #page1:nth-child(1) tr:nth-child(18) tr:last-child td")[3].innerText'))
FileOverwrite(Yaml(currencies), "currency.json")
chrome.CloseTab()