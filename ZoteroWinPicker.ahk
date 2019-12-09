#NoEnv
; Menu, Tray, Icon, ZoteroWinPicker.ico
#Persistent 
Menu, Tray, NoStandard 
Menu, Tray, Add, Zotero Picker Settings, openSettings  
Menu, Tray, Add, , ;
Menu, Tray, Add, About, about
Menu, Tray, Add
Menu, Tray, Add, Exit, exit

iniFileName = ZoteroWinPicker.ini
pickerWindowTitle = Quick Format Citation
formatOptions := {"Latex":"latex","Biblatex":"biblatex","MultiMarkdown":"mmd","Pandoc":"pandoc","Zotero ODF Scan":"scannable-cite","Formatted Zotero Quick Citation":"formatted-citation","Formatted Zotero Quick Bibliography":"formatted-bibliography","JSON":"json"}
odtLocatorOptions := {"Article":"art.","Chapter":"ch.","Subchapter":"subch.","Column":"col.","Figure":"fig.","Line":"l.","Note":"n.","Issue":"no.","Opus":"op.","Page":"p.","Paragraph":"para.","Subparagraph":"subpara.","Part":"pt.","Rule":"r.","Section":"sec.","Subsection":"subsec.","Section":"Sec.","Sub verbo":"sv.","Schedule":"sch.","Title":"tit.","Verse":"vrs.","Volume":"vol."}


IfNotExist, %iniFileName%
    {
    MsgBox, ,Zotero Windows Picker, 
    (
        Initialize Settings
    )
    currentFormat := "Zotero ODF Scan"
    currentShortcut := "^!F"
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
global tempClipboard := Clipboard
IfWinNotExist, ahk_exe zotero.exe
{
    MsgBox, , Zotero Not Found, Please Launch Zotero for Windows First.
    Return
}
WinGetActiveTitle, activeWindow
Clipboard := ""
req := ComObjCreate("WinHttp.WinHttpRequest.5.1")
req.Open("GET", requestString, false)
req.SetRequestHeader("Content-Type", "application/json")
try 
{
req.Send()
} catch {
    MsgBox, , Error,"Failed to Connect. Check the Settings for Zotero for Windows and Better BibTeX."
    WinClose, %pickerWindowTitle%
    return
}
IfWinExist, %pickerWindowTitle%
    WinActivate, %pickerWindowTitle%
    WinSet, AlwaysOnTop, ON, %pickerWindowTitle%

req.WaitForResponse(5)
if (!req.ResponseText){
    return
}
if (InStr(req.ResponseText, "CAYW failed: Error:")) {
    errorMessage := req.ResponseText 
    MsgBox,, Error,
    (
To use formatted citation, set Zotero default quick-copy format to a citation style first.

Error Message: %errorMessage%
    )
    return
}
citationStrings := req.ResponseText
if (currentFormat = "Zotero ODF Scan" && locatorCheck = 1) 
{
rList := formatResult(citationStrings)
for index, value in rList
    {
    Gui, 2: Add, Text, , %index%. %value% 
    Gui, 2: Add, DropDownList, vlocator%index% Choose9, %odtLocatorOptionsList%
    Gui, 2: Add, Edit,  vlocatorNumber%index% 
    }
Gui,2: Add, Button, default, OK  ; The label ButtonOK (if it exists) will be run when the button is pressed.
Gui,2: Show,, Locator Information
return  ; End of auto-execute section. The script is idle until the user does something.
2GuiClose:
Gui 2: Destroy
return
2ButtonOK:
Gui,2: Submit  ; Save the input from the user to each control's associated variable.
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
Gui,2: Destroy
outputResult(citationStrings)
return
} else {
outputResult(citationStrings)
}

outputResult(citationStrings) {
global
Clipboard := citationStrings
ClipWait
if (insertCheck = 1) 
    {
    WinActivate, %activeWindow%
    Sleep, 200
    Send, ^v
    if (notificationCheck = 1)
    {
        TrayTip, Citation Inserted, %citationStrings%, 4,
        Clipboard := tempClipboard
    }
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

exit() {
    ExitApp
}

About() {
    MsgBox, , Zotero Windows Picker 0.9,
    (
Requirement: 

    1. Zotero For Window
    2. Better BibTeX for Zotero

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
    ; return currentFormat,currentShortcut, locatorCheck, insertCheck
}

formatShortcut(s){
    s := StrReplace(s, "+", "Shift + ")
    s := StrReplace(s, "!", "Alt + ")
    s := StrReplace(s, "^", "Ctrl + ")
    return s
}

openSettings(){
    global
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
    Gui, Add, Checkbox, w65 x+25 vctrlCheck %checkCtrl% , Ctrl 
    
    Gui, Add, Checkbox, w65 x+1 vshiftCheck %checkShift% , Shift
    
    Gui, Add, Checkbox, w65 x+1 valtCheck %checkAlt% , Alt

    Gui, Add, Text, w45 x+1  , +

    Gui, Add, DropDownList, w90 x+1 vchosenLetter %checkKey% , %letterList%

    Gui, Add, Text, x10, 

    Gui, Add, Text, w75 x10, Format:

    Gui, Add, DropDownList, w180 x+25 vchosenFormat %checkFormat% , %formatOptionList%
    Gui, Add, Link, x+10, See <a href="https://retorque.re/zotero-better-bibtex/citing/cayw/">Better BibTex Documentation</a> for details.

    Gui, Add, Text, x10, 

    Gui, Add, Text, w75 x10, Zotero ODT Scan:
    Gui, Add, Checkbox, x+25 vlocatorCheck %checkLocator%, Input Locator Information (add informations like pages etc during insertion) 
    
    Gui, Add, Text, x10, 

    Gui, Add, Text, w75 x10, Export:
    Gui, Add, Checkbox, x+25 vinsertCheck %checkInsert%, Automatically Insert Citations (uncheck to save to clipboard only)

    Gui, Add, Text, x10, 

    Gui, Add, Text, w75 x10, Notification:
    Gui, Add, Checkbox, x+25 vnotificationCheck %checkNotification%, Show Notification;
    
    Gui, Add, Text, x10, 

    Gui, Add, Text, x10, * Zotero Format Picker window might hide in the background when called for the first time.

    Gui, Add, Link, x10, * For Zotero ODT Scan instructions, see: <a href="https://zotero-odf-scan.github.io/zotero-odf-scan/">here.</a> 

    Gui, Add, Text, x10, 


    Gui, Add, Button, w160 default, Save
    Gui, Add, Button, w140 x+10, Cancel



    Gui, Show,, Zotero Windows Picker Settings
    return  ; End of auto-execute section. The script is idle until the user does something.

    ButtonCancel:
    GuiClose:
    Gui, Destroy
    return
    ButtonSave:
    Gui, Submit  ; Save the input from the user to each control's associated variable.
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
    MsgBox, Settings Saved.
    Reload
}
