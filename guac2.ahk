/*	GUACAMOLE conference data browser (C)2023 TC
	Version 2 - uses AHK2 and improved GUI
*/
#Requires AutoHotkey v2.0
SetWorkingDir(A_ScriptDir)
#Include "%A_ScriptDir%\Includes"

Initialization:
{
	if WinExist("View Downloads - Windows Internet Explorer") {
		WinClose
	}

/*	Set environment and vars
*/
	user := A_UserName
	isDevt := InStr(A_WorkingDir,"AhkProjects")
	if (isDevt) {
		netdir := A_WorkingDir "\devfiles\Tuesday_Conference"						; local files
		chipdir := A_WorkingDir "\devfiles\CHIPOTLE\"
		confStart := "20220614140000"
	} else {
		netdir := "\\childrens\files\HCConference\Tuesday_Conference"				; networked Conference folder
		chipdir := "\\childrens\files\HCChipotle\"									; and CHIPOTLE files
		confStart := A_Now
	}
	epRead := readIni("epRead")
	fc := readIni("forecast")
}

; y := ComObject("Msxml2.DOMDocument.3.0")
; y.load("devfiles\guac.xml")

; k := y.selectNodes("/root/id")
; MsgBox k.Length

ExitApp


readIni(section) {
/*	Reads a set of variables

	[section]					==	 		var1 := res1, var2 := res2
	var1=res1
	var2=res2
	
	[array]						==			array := ["ccc","bbb","aaa"]
	=ccc
	=bbb
	=aaa
	
	[objet]						==	 		objet := {aaa:10,bbb:27,ccc:31}
	aaa:10
	bbb:27
	ccc:31
*/
	global
	local x, i, key, val, k, v
		, i_res
		, i_type := []
		, i_lines := []
		, iniFile := ".\files\guac.ini"
	i_type.var := i_type.obj := i_type.arr := false

	x:=IniRead(iniFile,section)
	loop parse x, "`n", "`r"
	{
		i := A_LoopField
		if (i~="(?<!`")[=]") 															; find = not preceded by "
		{
			if (i ~= "^=") {															; starts with "=" is an array list
				i_type.arr := true
				i_res := Array()
			} else {																	; "aaa=123" is a var declaration
				i_type.var := true
			}
		} 
		else																			; does not contain a quoted =
		{
			if (i~="(?<!`")[:]") {														; find : not preceded by " is an object
				i_type.obj := true
				i_res := Map()
		} else {																		; contains neither = nor : can be an array list
				i_type.arr := true
				i_res := Array()
			}
		}
	}
	if ((i_type.obj) + (i_type.arr) + (i_type.var)) > 1 {								; too many types, return error
		return error
	}
	Loop parse x, "`n","`r"																; now loop through lines
	{
		i := A_LoopField
		if (i_type.var) {
			key := strX(i,"",1,0,"=",1,1)
			val := trim(strX(i,"=",1,1,"",1,0),'`"')
			k := &key
			v := &val
			%k% := %v%
		}
		if (i_type.obj) {
			key := trim(strX(i,"",1,0,":",1,1),'`"')
			val := trim(strX(i,":",1,1,"",0),'`"')
			i_res[key] := val
		}
		if (i_type.arr) {
			i := RegExReplace(i,"^=")													; remove preceding =
			i_res.push(trim(i,'`"'))
		}
	}
	return i_res
}

;	============ INCLUDES =================
; #Include xml.ahk
#Include strx2.ahk
; #Include Class_LV_Colors.ahk
; #Include sift3.ahk
; #Include CMsgBox.ahk