/*	GUACAMOLE conference data browser (C)2023 TC
	Version 2 - uses AHK2 and improved GUI
*/
#Requires AutoHotkey v2.0
SetWorkingDir(A_ScriptDir)
#Include "%A_ScriptDir%\Includes"

if WinExist("View Downloads - Windows Internet Explorer") {
	WinClose
}


;	============ INCLUDES =================
; #Include xml.ahk
#Include strx2.ahk
; #Include Class_LV_Colors.ahk
; #Include sift3.ahk
; #Include CMsgBox.ahk