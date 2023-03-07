/*	GUACAMOLE conference data browser (C)2023 TC
	Version 2 - uses AHK2 and improved GUI
*/
#Requires AutoHotkey v2.0
SetWorkingDir(A_ScriptDir)
#Include "%A_ScriptDir%\Includes"
#Warn VarUnset, OutputDebug

;#region == Initialization ==============================================================================
/*	Set environment and vars
*/
	user := A_UserName
	isDevt := InStr(A_WorkingDir,"AhkProjects")
	if (isDevt) {
		netdir := A_WorkingDir "\devfiles\Tuesday_Conference"							; local files
		chipdir := A_WorkingDir "\devfiles\CHIPOTLE\"
		confStart := "20220614140000"
	} else {
		netdir := "\\childrens\files\HCConference\Tuesday_Conference"					; networked Conference folder
		chipdir := "\\childrens\files\HCChipotle\"										; and CHIPOTLE files
		confStart := A_Now
	}
	res := MsgBox("Are you launching GUACAMOLE for patient presentation?","GUACAMOLE",36)
	if res="Yes"
		isPresenter := true
	else
		isPresenter := false

	firstRun := true
	; SplashImage, % chipDir "guac.jpg", B2 

	y := ComObject("Msxml2.DOMDocument")
	y.load(chipdir "currlist.xml")														; Get latest local currlist into memory

	arch := ComObject("Msxml2.DOMDocument")
	arch.load(chipdir "archlist.xml")													; Get archive.xml

	datedir := Map()
	datedir.Default := Map()
	mo := ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

	confList := Map()

	screen := {guiW:1200,guiH:400,Width:A_ScreenWidth,Height:A_ScreenHeight}

	; RegCOM(".\Includes\dsoframer.manifest")
;#endregion

;#region == Main Loop ===================================================================================
	MainGUI()																			; Draw the main GUI
	if (firstRun) {
		;~ SoundPlay, % chipDir "chillin.wav", Wait
		; SplashImage, off
		firstRun := false
	}
	SetTimer(confTimer, 1000)															; Update ConfTime every 1000 ms
	; WinWaitClose, GUACAMOLE Main														; wait until main GUI is closed

ExitApp
;#endregion

;#region == TIMERS ===============================================================================================
confTimer() {
	global isPresenter, confStart
	tmp := FormatTime(A_now,"HH:mm:ss")													; Format the current time
	; GuiControl, mainUI:Text, CTime, % tmp												; Update the main GUI current time
	
	if (isPresenter) {																	; For presenter only,
		tt := elapsed(confStart,A_Now)													; Total time elapsed
		; GuiControl, mainUI:Text, CDur, % tt.HHMMSS									; Update the main GUI elapsed time
	}
	Return
}

elapsed(start,end) {
	tdiff := formatSec(DateDiff(end,start,"Seconds"))
	return tdiff
}

formatSec(secs) {
	HH := zDigit(floor(secs/3600))														; Derive HH from total time (secs)
	MM := zDigit(floor((secs-HH*3600)/60))												; Derive MM from remainder of HH
	SS := zDigit(secs-HH*3600-MM*60)													; Derive SS from remainder of MM
	Return {hh:HH, mm:MM, ss:SS
		, HHMMSS:HH ":" MM ":" SS
		, HHMM:HH ":" MM
		, MMSS:MM ":" SS}
}
;#endregion  ============================================================================================

;#region == MAIN GUI ===============================================================================================
MainGUI()
{
	global confDate, isDevt

	if !IsSet(confDate) {
		if (isDevt) {
			confDate := GetConfDate("20220614")											; use test dir. change this if want "live" handling
		} else {
			confDate := GetConfDate()													; determine next conference date into array dt
		}
	}
	GetConfDir(confDate)																; find confList, confXls, gXml

	; Gui, mainUI:Default
	; Gui, Destroy
	; Gui, Font, s16 wBold
	; Gui, Add, Text, y26 x20 vCTime, % "              "								; Conference real time
	; Gui, Add, Text, % "y26 x" winDim.gw-100 " vCDur", % "              "			; Conference duration (only exists for Presenter)
	; Gui, Add, Text, % "y0 x0 w" winDim.gw " h20 +Center", .-= GUACAMOLE =-.
	; Gui, Font, wNorm s8 wItalic
	; Gui, Add, Text, yp+30 xp wp +Center, General Use Access for Conference Archive
	; Gui, Add, Text, yp+14 xp wp +Center, Merged OnLine Elements
	; Gui, Add, Text, y10 x54, Time
	; Gui, Add, Text, % "y10 x" winDim.gw-72, Duration
	; Gui, Font, wBold
	; Gui, Font, wNorm
	; makeConfLV()																	; Draw the pateint grid ListView
	; Gui, Add, Button, wp +Center gDateGUI, % confDate.MDY							; Date selector button
	; Gui, Show, AutoSize, % "GUACAMOLE Main - " confDate.MDY							; Show GUI with seleted conference DT
	Return
}
;#endregion  ============================================================================================

;#region == CONFERENCE DIRECTORIES ======================================================================
GetConfDate(dt:=A_Now) {
; Get next conference date. If not argument, assume today
	Wday := FormatTime(dt,"WDay")														; Today's day of the week (Sun=1)
	if (Wday > 3) {																		; if Wed-Sat, next Tue
		t2 := DateAdd(dt,10-Wday,"Days")
	} else {																			; if Sun-Tue, this Tue
		t2 := DateAdd(dt,3-Wday,"Days")
	}
	conf := ParseDate(dt)
	return {YYYY:conf.YYYY, MM:conf.MM, MMM:conf.MMM, DD:conf.DD
		, YMD:conf.YMD, MDY:conf.MDY}
}

GetConfDir(confDate) {
/*	Find conference folder path for confDate
	Get list of patient folders, push to confList, save in guac.xml
*/
	global confList, confXls, firstRun, gXml, confDir

	confDir := NetConfDir(confDate.YYYY,confDate.mmm,confDate.dd)						; get path to conference folder based on predicted date "confDate"
	SetWorkingDir(netdir "\" confDir)
	; if !IsObject(confList) {															; make sure confList array exists
	; 	confList := {}
	; }
	If FileExist("guac.xml")
	{
		gXml := ComObject("Msxml2.DOMDocument")
		gXml.load("guac.xml")															; Open existing guac.xml
	} else {
		gXml := ComObject("Msxml2.DOMDocument")
		gXml.loadXML("<root/>")															; Create new blank guac.xml if it doesn't exist
		gXml.save("guac.xml")
	}
	filelist := ""																		; Clear out filelist string
	patnum := ""																		; and zero out count of patient folders
	
	; Progress,,,Reading conference directory
	Loop Files ".\*", "DF"																; Loop through all files and directories in confDir
	{
		tmpNm := A_LoopFileName
		tmpExt := A_LoopFileExt
		if (tmpNm ~= "i)Fast.Track|-FT|\sFT")											; exclude Fast Track files and folders
			continue
		if (tmpExt) {																	; evaluate files with extensions
			if (tmpNm ~= "i)(\~\$|(Fast.Track|-FT|\sFT))")								; exclude temp and "Fast Track" files
				continue
			if (tmpNm ~= "i)(PCC)?.*\d{1,2}\.\d{1,2}\.\d{2,4}.*xls") {					; find XLS that matches PCC 3.29.16.xlsx
				confXls := tmpNm
			}
			continue
		}
		tmpNm := RegExReplace(tmpNm,"\'","_")
		if !(confList.Has(tmpNm)) {														; confList is empty
			tmpNmUP := RegExReplace(format("{:U}",tmpNm),"\'","_")						; place filename in all UPPER CASE
			confList[tmpNmUP] := {name:tmpNm,done:0,note:""}							; name=actual filename, done=no, note=cleared
		}
		tmpPath := "/root/id[@name='" tmpNmUP "']"
		if !IsObject(gXml.selectSingleNode(tmpPath)) {
			e := gXml.createElement("id")
			e.setAttribute("name",tmpNmUP)
			gXml.appendChild(e)
			; gXml.addElement("id","root",{name: tmpNmUP})					; Add to Guac XML if not present
		}
	}
	; if (confXls) {															; Read confXls if present
	; 	Progress, % (firstRun)?"off":"",,Reading XLS file
	; 	readXls()
	; }
	; gXml.save("guac.xml")													; Write Guac XML
	; Return
}

makeConfLV() {
	; global confList, winDim, gXml, mainUI

	; Gui, mainUI:Default
	; Gui, Font, s16
	; Gui, Add, ListView, % "r" confList.length()+1 " x20 w" windim.gw-20
	; 	. " Hdr AltSubmit Grid BackgroundSilver NoSortHdr NoSort gPatDir"
	; 	, Name|Done|Takt|Diagnosis|Note
	; Progress, % (firstRun)?"off":"",,Rendering conference list
	; for key,val in confList
	; {
	; 	if (key=A_index) {
	; 		keyNm := confList[key]											; UPPER CASE name
	; 		keyElement := "/root/id[@name='" keyNm "']"
	; 		keyDx := (tmp:=gXml.selectSingleNode(keyElement "/diagnosis").text) ? tmp : ""	; DIAGNOSIS, if present
	; 		keyDone := gXml.getAtt(keyElement,"done")						; DONE flag
	; 		keyDur := (tmp:=gXml.getAtt(keyElement,"dur")) ? formatSec(tmp) : ""	; If dur exists, get it
	; 		keyNote := (tmp:=gXml.selectSingleNode(keyElement "/notes").text) ? tmp : ""	; NOTE, if present
	; 		LV_Add(""
	; 			,keyNm														; UPPER CASE name
	; 			,(keyDone) ? "x" : ""										; DONE or not
	; 			,(keyDur) ? keyDur.MM ":" keyDur.SS : ""					; total DUR spent on this patient MM:SS
	; 			,(keyDx) ? keyDx : ""										; Diagnosis
	; 			,(keyNote) ? keyNote : "")									; note for this patient
	; 	}
	; }
	; Progress, Off
	; LV_ModifyCol()
	; LV_ModifyCol(1,"200")
	; LV_ModifyCol(2,"AutoHdr Center")
	; LV_ModifyCol(3,"AutoHdr Center")
	; LV_ModifyCol(4,"AutoHdr")
	; LV_ModifyCol(5,"AutoHdr")
	; Return
}

NetConfDir(yyyy:="",mmm:="",dd:="") {
	global netdir, datedir, mo

	
	if (datedir[yyyy].Has(mmm)) {														; YYYY\MMM already exists
		return yyyy "\" datedir[yyyy,mmm].dir "\" datedir[yyyy,mmm,dd]					; return the string for YYYY\MMM
	}
	if !(datedir.Has(yyyy)) {
		datedir[yyyy] := Map()
	}
	Loop Files netdir "\" yyyy "\*", "D"												; Get the month dirs in YYYY
	{
		file := A_LoopFileName
		for key,obj in mo																; Compare "file" name with Mo abbrevs
		{
			if (instr(file,obj)) {														; mo MMM abbrev is in A_loopfilename
				datedir[yyyy][obj] := Map()
				datedir[yyyy][obj].dir := file 											; insert wonky name as yr[yyyy,mmm,{dir:filename}]
			}
		}
	}
	Loop Files netdir "\" yyyy "\" datedir[yyyy][mmm].dir "\*" , "D"					; check for conf dates within that month (dir:filename)
	{
		file := A_LoopFileName
		if (regexmatch(file,"\d{1,2}\.\d{1,2}\.\d{1,2}")) {								; sometimes named "6.19.15"
			d0 := zdigit(strX(file,".",1,1,".",1,1))
			datedir[yyyy][mmm][d0] := file
		} else if (RegExMatch(file,"\w\s\d{1,2}")){										; sometimes named "Jun 19" or "June 19"
			d0 := zdigit(strX(file," ",1,1,"",1,0))
			datedir[yyyy][mmm][d0] := file
		} else if (regexmatch(file,"\b\d{1,2}\b")) {									; sometimes just named "19"
			d0 := zdigit(file)
			datedir[yyyy][mmm][d0] := file
		}																				; inserts dir name into datedir[yyyy,mmm,dd]
	}
	return yyyy "\" datedir[yyyy][mmm].dir "\" datedir[yyyy][mmm][dd]						; returns path to that date's conference 
}

	
;#endregion

;#region == FORMATTING =====================================================================================
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

zDigit(x) {
; Returns 2 digit number with leading 0
	return SubStr("00" x, -2)
}

strQ(var1,txt,null:="") {
/*	Print Query - Returns text based on presence of var
	var1	= var to query
	txt		= text to return with ### on spot to insert var1 if present
	null	= text to return if var1="", defaults to ""
*/
	return (var1="") ? null : RegExReplace(txt,"###",var1)
}

ObjHasValue(aObj, aValue, rx:="") {
	for key, val in aObj
		if (rx="RX") {																	; argument 3 is "RX" 
			if (aValue="") {															; null aValue in "RX" is error
				return false
			}
			if (val ~= aValue) {														; val=text, aValue=RX
				return key
			}
			if (aValue ~= val) {														; aValue=text, val=RX
				return key
			}
		} else {
			if (val = aValue) {															; otherwise just string match
				return key
			}
		}
	return false																		; fails match, return err
}

;#endregion

; ============ INCLUDES =================
; #Include xml.ahk
#Include strx2.ahk
; #Include MsgBox2.ahk
; #Include Class_LV_Colors.ahk
; #Include sift3.ahk
; #Include CMsgBox.ahk