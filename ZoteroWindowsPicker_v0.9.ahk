#SingleInstance, Force
#NoEnv
; Menu, Tray, Icon, ZoteroWinPicker.ico
#Persistent 
Menu, Tray, NoStandard 
Menu, Tray, Add, Zotero Picker Settings, openSettings  
Menu, Tray, Add
Menu, Tray, Add, Restart, reload
Menu, Tray, Add 
Menu, Tray, Add, About, about
Menu, Tray, Add
Menu, Tray, Add, Exit, exit


iniFileName = ZoteroWindowsPicker.ini
pickerWindowTitle = Quick Format Citation
formatOptions := {"Latex":"latex","Biblatex":"biblatex","MultiMarkdown":"mmd","Pandoc":"pandoc","Zotero ODF Scan":"scannable-cite","Formatted Zotero Quick Citation":"formatted-citation","Formatted Zotero Quick Bibliography":"formatted-bibliography","JSON":"json"}
odtLocatorOptions := {"Article":"art.","Chapter":"ch.","Subchapter":"subch.","Column":"col.","Figure":"fig.","Line":"l.","Note":"n.","Issue":"no.","Opus":"op.","Page":"p.","Paragraph":"para.","Subparagraph":"subpara.","Part":"pt.","Rule":"r.","Section":"sec.","Subsection":"subsec.","Section":"Sec.","Sub verbo":"sv.","Schedule":"sch.","Title":"tit.","Verse":"vrs.","Volume":"vol."}


IfNotExist, %iniFileName%
    {
    MsgBox, ,Zotero Windows Picker, 
    (
        Initialial Settings
    )
    currentFormat := "Formatted Zotero Quick Bibliography"
    currentShortcut := "+!F"
    locatorCheck := 1
    insertCheck := 1
    notificationCheck := 1
    openSettings()
    } else {
    readIni()
    }

if (currentFormat = "Zotero ODF Scan")
{
for key, value in odtLocatorOptions
{
    odtLocatorOptionsList .= key . "|"
}
}
formatedShortcut := formatShortcut(currentShortcut)
Menu, Tray, Tip , Format:`n     %currentFormat%`n`nShortcut:`n      %formatedShortcut%`n

currentFormatString := "format=" . formatOptions[currentFormat]
requestString := "http://127.0.0.1:23119/better-bibtex/cayw?" . currentFormatString
Hotkey, %currentShortcut%, getRef
return

getRef:
IfWinExist, Quick Format Citation
WinActivate, Quick Format Citation
global tempClipboard := Clipboard
Clipboard := ""
IfWinNotExist, ahk_exe zotero.exe
{
    MsgBox, , Zotero Not Found, Please Launch Zotero for Windows First.
    Return
}
WinGetActiveTitle, activeWindow
req := ComObjCreate("WinHttp.WinHttpRequest.5.1")
; req := ComObjCreate("Msxml2.ServerXMLHTTP")
req.Open("GET", requestString, true)
req.SetRequestHeader("Content-Type", "application/json")
try 
{    
req.Send()
req.WaitForResponse()
} catch {
    WinClose, %pickerWindowTitle%
    return
}
if (!req.ResponseText){
    return
}
if (InStr(req.ResponseText, "CAYW failed: Error: scannable-cite")) {
    errorMessage := req.ResponseText 
    MsgBox,, Remainder,
    (

To use ODF format, you need to install "RTF/ODF Scan for Zotero" plugin.

See: https://zotero-odf-scan.github.io/zotero-odf-scan/

To switch to other formats, go to Settings.


Error Message: %errorMessage%

    )
    return
} else if (InStr(req.ResponseText, "CAYW failed: Error: formatted")) {
    errorMessage := req.ResponseText 
    MsgBox,, Remainder,
    (

To use formatted citation, set Zotero default quick-copy format to a citation style first.

If you want to switch to other formats, go to Settings.


Error Message: %errorMessage%

    )
    return
}
else if (InStr(req.ResponseText, "No endpoint")) {
    errorMessage := req.ResponseText 
    MsgBox,, Remainder,
    (

Cannot connect to Better BibTeX for Zotero add-on. Install it and restart Zotero.

See: https://retorque.re/zotero-better-bibtex/


Error Message: %errorMessage%

    )
    return
}
else if (InStr(req.ResponseText, "CAYW failed: translation")) {
    errorMessage := req.ResponseText 
    MsgBox,, Remainder,
    (

Cannot find ODT translator. Check or reinstall Zotero ODT Scan add-on.

See: https://zotero-odf-scan.github.io/zotero-odf-scan/


Error Message: %errorMessage%

    )
    return
}
citationStrings := req.ResponseText
if (currentFormat = "Zotero ODF Scan" && locatorCheck = 1) 
{
IfWinExist, Locator Information
Gui,2:Destroy
rList := formatResult(citationStrings)
for index, value in rList
    {
    Gui, 2: Add, Text, , %index%. %value% 
    Gui, 2: Add, DropDownList, vlocator%index% Choose9, %odtLocatorOptionsList%
    Gui, 2: Add, Edit,  vlocatorNumber%index% 
    }
Gui,2: Add, Button, w120 default, OK 
Gui,2: Add, Button, w120 x+5 , Cancel 
Gui,2: Show,, Locator Information
return  
2GuiClose:
2ButtonCancel:
Gui 2: Destroy
return
2ButtonOK:
Gui,2: Submit  
Gui 2: Destroy
pList := []
for index, value in rList
    {
    locatorString := ""
    locator := "locator" index
    locator := %locator%
    ln := "locatorNumber" index    
    ln := %ln%
    if (ln){
        locatorString := "|" . " " . odtLocatorOptions[locator] . " " . ln
        value := RegExReplace(value, "\|", locatorString,, 1, InStr(value, "|",,, 2))
        pList.Push(value)
    }
    }
citationStrings := ""
for key, value in pList
    {
    citationStrings .= value
    }
outputResult(citationStrings)
return
} else {
outputResult(citationStrings)
return
}
return

outputResult(citationStrings) {
global
Clipboard := ""
Clipboard := citationStrings
ClipWait
; WinActivate, %activeWindow%
; WinWaitActive, %activeWindow%
if (insertCheck = 1) 
    {
    if (notificationCheck = 1)
    {
        TrayTip, Citation Inserted, %citationStrings%, 4,

    }
    Sleep, 500
    Send, ^v
    WinWait A
    Clipboard := tempClipboard
}else {
    if (notificationCheck = 1)
    {
    TrayTip, Saved to Clipboard, %citationStrings%, 4,
    }
}
}
return

formatResult(r) {
    r := StrReplace(r,"}{","};{")
    rList := StrSplit(r,";")
    return rList
}
return

About() {
    MsgBox, , Zotero Windows Picker 0.9,
    (
Requirement: 

    1. Zotero For Window
    2. Better BibTeX for Zotero
    3. RTF/ODF Scan for Zotero (for ODF format, optional)


Bo An
2019
Scripted in AHK.
    )
}

readIni() {
    global
    IniRead, currentFormat, %iniFileName%, Settings, currentFormat
    IniRead, currentShortcut, %iniFileName%, Settings, currentShortcut
    IniRead, locatorCheck, %iniFileName%, Settings, locatorCheck
    IniRead, insertCheck, %iniFileName%, Settings, insertCheck
    IniRead, notificationCheck, %iniFileName%, Settings, notificationCheck    
}

formatShortcut(s){
    s := StrReplace(s, "+", "Shift + ")
    s := StrReplace(s, "!", "Alt + ")
    s := StrReplace(s, "^", "Ctrl + ")
    return s
}

openSettings(){
    global
    IfWinExist, Zotero Windows Picker Settings
    {
        WinActivate, Zotero Windows Picker Settings
        return
    }
    Start := A_TickCount
    loop, 26
    {
    letterList .= chr(A_Index + 64) . "|"
    }
    loop, 10
    {
    letterList .= chr(A_Index + 47) . "|"
    }
    loop, 12
    {
    letterList .= "F" . A_Index . "|"
    }

    for key, value in formatOptions
    {
        formatOptionList .= key . "|"
    }
    IfExist, %iniFileName% 
        readIni()
    checkCtrl := InStr(currentShortcut, "^") ? "Checked" : ""
    checkAlt := InStr(currentShortcut, "!") ? "Checked" : ""
    checkShift := InStr(currentShortcut, "+") ? "Checked" : ""
    currentKeyChoice := RegExReplace(currentShortcut,"[!+^]")
    checkLocator := locatorCheck ? "Checked" : ""
    checkInsert := insertCheck ? "Checked" : ""
    checkNotification := notificationCheck ? "Checked": ""

    letterListArray := StrSplit(letterList,"|")
    formatOptionListArray := StrSplit(formatOptionList,"|")
    for index, value in letterListArray
        if (value = currentKeyChoice)
            checkKey := "Choose" . index
    for index, value in formatOptionListArray
        if (value = currentFormat)
            checkFormat := "Choose" . index

    ; GUI for Settings
    Gui, Add, Text, w75 y10  , Hotkey:
    Gui, Add, Checkbox, w65 x+15 vctrlCheck %checkCtrl% , Ctrl 
    
    Gui, Add, Checkbox, w65 x+1 vshiftCheck %checkShift% , Shift
    
    Gui, Add, Checkbox, w65 x+1 valtCheck %checkAlt% , Alt

    Gui, Add, Text, w45 x+1  , +

    Gui, Add, DropDownList, w90 x+1 vchosenLetter %checkKey% , %letterList%

    Gui, Add, Text, x10, 

    Gui, Add, Text, w75 x10, Citation Format:

    Gui, Add, DropDownList, w200 x+25 vchosenFormat %checkFormat% , %formatOptionList%

    Gui, Add, Link, x110, * Use the ODT format if you want link and update your citations with Zotero. 
    Gui, Add, Link, x110, * For details of different formats, see <a href="https://retorque.re/zotero-better-bibtex/citing/cayw/">Better BibTex documentation</a>.

    Gui, Add, Text, x10, 

    Gui, Add, Text, w75 x10, Zotero ODT Scan:
    Gui, Add, Checkbox, x+15 vlocatorCheck %checkLocator%, Input Locator Information (add informations like pages etc during insertion) 
    Gui, Add, Link, x110, * For detailed Zotero ODT Scan instructions, see <a href="https://zotero-odf-scan.github.io/zotero-odf-scan/">the Add-on website</a> 

    Gui, Add, Text, x10, 

    Gui, Add, Text, w75 x10, Export:
    Gui, Add, Checkbox, x+15 vinsertCheck %checkInsert%, Automatically Insert Citations (uncheck to save to clipboard only)

    Gui, Add, Text, x10, 

    Gui, Add, Text, w75 x10, Notification:
    Gui, Add, Checkbox, x+15 vnotificationCheck %checkNotification%, Show Notification
    
    Gui, Add, Text, x10, 

    Gui, Add, Button, w160 default, Save
    Gui, Add, Button, w140 x+10, Cancel
    Gui, Add, Link, x+10 w130 Right, v0.9 by Bo An via <a href="https://www.autohotkey.com">AHK</a>.


    Gui, Show,, Zotero Windows Picker Settings
    return  

    ButtonCancel:
    GuiClose:
    Gui, Destroy
    return
    ButtonSave:
    Gui, Submit  
    if(ctrlCheck){
        chosenShortcut .= "^"
    }
    if(shiftCheck){
        chosenShortcut .= "+"
    }
    if(altCheck){
        chosenShortcut .= "!"
    }
    chosenShortcut .= chosenLetter
    IniWrite, %chosenFormat%, %iniFileName%, Settings, currentFormat
    IniWrite, %chosenShortcut%, %iniFileName%, Settings, currentShortcut
    IniWrite, %locatorCheck%, %iniFileName%, Settings, locatorCheck
    IniWrite, %insertCheck%, %iniFileName%, Settings, insertCheck
    IniWrite, %notificationCheck%, %iniFileName%, Settings, notificationCheck
    MsgBox, Settings Saved
    Reload
}

OnWin(Event, Hwnd)	
{
Static This_Func_Name := "OnWin"
Static RunAtScriptExecution1 := DllCall( "RegisterShellHookWindow", UInt, A_ScriptHwnd)		
Static SH_MsgNum := DllCall( "RegisterWindowMessage", Str,"SHELLHOOK" )
Static RunAtScriptExecution2 := OnMessage(SH_MsgNum, Func(This_Func_Name), 1000)
	; ((event = 32772) || (event = 4))
    if (event = 1)
	{
	WinGetTitle, Title_Found, % "ahk_id" Hwnd	
	
	if (Title_Found = "Quick Format Citation")
		{
        IfWinNotActive, Quick Format Citation
        {
        WinActivate, Quick Format Citation
        }
        }
	}
}

reload(){
    reload
}

exit() {
    ExitApp
}