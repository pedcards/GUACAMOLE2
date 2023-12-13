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
		confStart := "20220614140000"
	} else {
		netdir := "\\childrens\files\HCConference\Tuesday_Conference"					; networked Conference folder
		confStart := A_Now
	}
	res := MsgBox("Are you launching GUACAMOLE for patient presentation?","GUACAMOLE",36)
	if res="Yes"
		isPresenter := true
	else
		isPresenter := false

	firstRun := true
	; SplashImage, % netDir "\guac.jpg", B2 

	datedir := Map()
	datedir.Default := Map()
	mo := ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
	winDim := {gw:1200,gh:400,scrX:A_ScreenWidth,scrY:A_ScreenHeight}

	confList := Map()

	screen := {guiW:1200,guiH:400,Width:A_ScreenWidth,Height:A_ScreenHeight}

	; RegCOM(".\Includes\dsoframer.manifest")
;#endregion

;#region == Main Loop ===================================================================================
	MainGUI()																			; Draw the main GUI
	if (firstRun) {
		; SplashImage, off
		firstRun := false
	}
	SetTimer(confTimer, 1000)															; Update ConfTime every 1000 ms
	WinWaitClose("GUACAMOLE Main")														; wait until main GUI is closed

ExitApp()
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
	global confDate, isDevt, winDim, mainUI

	if !IsSet(confDate) {
		if (isDevt) {
			confDate := GetConfDate("20220614")											; use test dir. change this if want "live" handling
		} else {
			confDate := GetConfDate()													; determine next conference date into array dt
		}
	}
	GetConfDir(confDate)																; find confList, confXls, gXml

	mainUI := Gui("","GUACAMOLE Main - " confDate.MDY)
	mainUI.SetFont("s16 Bold")
	mainUI.Add("Text","y26 x20 vCTime","              ")								; Conference real time
	mainUI.Add("Text","y26 x" winDim.gw-100 " vCDur","              ")					; Conference duration (only exists for Presenter)
	mainUI.Add("Text","y0 x0 w" winDim.gw " h20 +Center",".-= GUACAMOLE =-.")
	mainUI.SetFont("s8 Norm Italic")
	mainUI.Add("Text","yp+30 xp wp +Center","General Use Access for Conference Archive")
	mainUI.Add("Text","yp+14 xp wp +Center","Merged OnLine Elements")
	mainUI.Add("Text","y10 x54","Time")
	mainUI.Add("Text","y10 x" winDim.gw-72,"Duration")
	mainUI.SetFont("Norm s16")
	makeConfLV()																		; Draw the patient grid ListView
	mainDateBtn := mainUI.Add("Button","wp +Center ", confDate.MDY)						; Date selector button
	mainDateBtn.OnEvent("Click",DateGUI)
	mainUI.Show("AutoSize")

	Return
}

DateGUI(*) {
/*	modified from https://www.autohotkey.com/boards/viewtopic.php?f=82&t=67841&p=299272&hilit=monthcal#p534601
*/
	global confDate
	dateGUI := Gui("AlwaysOnTop -MaximizeBox -MinimizeBox","Select PCC date...")
	dateGUI.Add("MonthCal","v_selectedDT",confDate.YMD).OnEvent("Change",f_updateDT)	; Show selected date and month selector
	; dateGUI.OnEvent("Escape",f_submitDT)
	dateSubmit := dateGUI.Add("Button","Center","Jump to date")
	dateSubmit.OnEvent("Click",f_submitDT)
	dateGUI.Show("AutoSize")
	WinWaitClose("Select PCC date")
	enc := s_updatedDT
	return

	f_updateDT(s_selectedDT,*) {
		s_updatedDT := s_selectedDT.value
	}
	f_submitDT(*) {
		dateGUI.Destroy()
	}

	; Gui, date:Destroy																; Close MonthCal UI
	; dt := GetConfDate(EncDt)														; Reacquire DT based on value
	; conflist =																		; Clear out confList
	; Gosub MainGUI																	; Redraw MainGUI
	; return
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

	gXml := ComObject("Msxml2.DOMDocument.6.0")
	If FileExist("guac.xml")
	{
		gXml.load("guac.xml")															; Open existing guac.xml
	} else {
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
		if (tmpNm ~= "i)Fast.?Track|-FT|\sFT|\sPrep\.")									; exclude Fast Track files and folders
			continue
		if (tmpExt) {																	; evaluate files with extensions
			if (tmpNm ~= "i)(\~\$|(Fast.?Track|-FT|\sFT|\sPrep\.))")								; exclude temp and "Fast Track" files
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
		if !IsObject(tmpE := gXml.selectSingleNode(tmpPath)) {
			xml.addElement(gXml.selectSingleNode("root"),"id",{name:tmpNmUP})			; Add to Guac XML if not present
		}
	}
	if (confXls) {																		; Read confXls if present
		; Progress, % (firstRun)?"off":"",,Reading XLS file
		readXls()
	}
	gXml.save("guac.xml")																; Write Guac XML
	Return
}

makeConfLV() {
	global confList, winDim, gXml, mainUI

	mainLV := mainUI.Add("ListView"
		, "r" confList.Count+1 " x20 w" windim.gw-20
		 	. " Hdr AltSubmit Grid BackgroundSilver NoSortHdr NoSort"
		, ["Name","Done","Takt","Diagnosis","Note"])
	mainLV.OnEvent("DoubleClick",PatDir)

	for name,val in confList
	{
		keyElement := "/root/id[@name='" name "']"
		keyNode := gXml.selectSingleNode(keyElement)
		keyDx := (tmp:=keyNode.selectSingleNode("diagnosis").text) ? tmp : ""			; DIAGNOSIS, if present
		keyDone := keyNode.getAttribute("done")											; DONE flag
		keyDur := (tmp:=keyNode.getAttribute("dur")) ? formatSec(tmp) : ""				; DUR, if present
		keyNote := (tmp:=keyNode.selectSingleNode("notes").text) ? tmp : ""				; NOTE, if present
		mainLV.Add(""
			,name														; UPPER CASE name
			,(keyDone) ? "x" : ""										; DONE or not
			,(keyDur) ? keyDur.MM ":" keyDur.SS : ""					; total DUR spent on this patient MM:SS
			,(keyDx) ? keyDx : ""										; Diagnosis
			,(keyNote) ? keyNote : "")									; note for this patient
	}
	; Progress, Off
	mainLV.ModifyCol()
	mainLV.ModifyCol(1,"200")
	mainLV.ModifyCol(2,"AutoHdr Center")
	mainLV.ModifyCol(3,"AutoHdr Center")
	mainLV.ModifyCol(4,"AutoHdr")
	mainLV.ModifyCol(5,"AutoHdr")
	
	Return
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
	return yyyy "\" datedir[yyyy][mmm].dir "\" datedir[yyyy][mmm][dd]					; returns path to that date's conference 
}

ReadXls() {
	global gXml, confXls

	if IsObject(tmpXml:=gXml.selectSingleNode("/root/done")) {							; last time ReadXLS run
		tmpXml := tmpXml.text
	}
	tmpXls:=FileGetTime(confXls)														; get XLS modified time
	if (DateDiff(tmpXls,tmpXml,"Seconds") < 0) {										; Compare XLS-XML time diff
		return
	}
 	FileCopy(confXls, "guac.xlsx", 1)													; Create a copy of the active XLS file 
	oWorkbook := ComObjGet(netDir "\" confDir "\guac.xlsx")								; Open the copy in memory (this is a one-way street)
	colArr := ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q"] 	; array of column letters
	xls_hdr := Map()
	xls_cel := Map()
	maxcol := 0
	Loop 
	{
		RowNum := A_Index																; Loop through rows in RowNum
		chk := oWorkbook.Sheets(1).Range("A" RowNum).value								; Get value in first column (e.g. A1..A10)
		if (RowNum=1) {																	; First row is just last file update
			upDate := chk
			continue
		}
		if !(chk)																		; if empty, then end of file or bad file
			break
		Loop
		{	
			ColNum := A_Index															; Iterate through columns
			if (colnum>maxcol)															; Extend maxcol (largest col) when we have passed the old max
				maxcol:=colnum
			cel := oWorkbook.Sheets(1).Range(colArr[ColNum] RowNum).value				; Get value of colNum rowNum (e.g. C4)
			if ((cel="") && (colnum=maxcol))											; Find max column
				break
			if (rownum=2) {																; Row 2 is headers
				; Patient name / MRN / Cardiologist / Diagnosis / conference prep / scheduling notes / presented / deferred / imaging needed / ICU LOS / Total LOS / Surgeons / time
				if instr(cel,"Patient name") {											; Fix some header names
					cel:="Name"
				}
				if instr(cel,"Conference prep") {
					cel:="Prep"
				}
				if instr(cel,"scheduling notes") {
					cel:="Notes"
				}
				if instr(cel,"imaging needed") {
					cel:="Imaging"
				}
				xls_hdr[ColNum] := trim(cel)											; Add cel to headers xls_hdr[]
			} else {
				xls_cel[ColNum] := cel													; Otherwise add value to xls_cel[]
			}
		}
		if (rownum=2) {
			continue
		}
		xls_mrn := Round(xls_cel[ObjHasValue(xls_hdr,"MRN")])							; Get value in xls_hdr MRN column 
		xls_name := xls_cel[ObjHasValue(xls_hdr,"Name")]								; Get name from xls_hdr Name column
		if !(xls_mrn)																	; Empty MRN, move on
			continue
		xls_nameL := RegExReplace(strX(xls_name,"",1,1,",",1,1),"\'","_")
		xls_nameUP := StrUpper(xls_nameL)												; Name in upper case
		xls_id := "/root/id[@name='" xls_nameUP "']"									; Element string for id[@name]
		
		if !IsObject(gXml.selectSingleNode(xls_id)) {									; Add new element if not present
			xml.addElement(gXml.selectSingleNode("root"),"id",{name:xls_nameUP})
		}
		gXlsID := gXml.selectSingleNode(xls_id)
		gXlsID.setAttribute("mrn",xls_mrn)												; Set MRN
		if !IsObject(gXlsID.selectSingleNode("name_full")) {							; Add full name if not present
			xml.addElement(gXlsID,"name_full",xls_name)
		}
		if !IsObject(gXlsID.selectSingleNode("diagnosis")) {							; Add diagnostics and Diagnosis
			xml.addElement(gXlsID,"diagnosis",xls_cel[ObjHasValue(xls_hdr,"Diagnosis")])
		}
		if !IsObject(gXlsID.selectSingleNode("prep")) {
			xml.addElement(gXlsID,"prep",xls_cel[ObjHasValue(xls_hdr,"Prep")])
		}
		if !IsObject(gXlsID.selectSingleNode("notes")) {
			xml.addElement(gXlsID,"notes",xls_cel[ObjHasValue(xls_hdr,"Notes")])
		}
	}
	if !IsObject(gXml.selectSingleNode("/root/done")) {
		xml.addElement(gXml.selectSingleNode("root"),"done",A_Now)						; Add <done> element when has been scanned to prevent future scans
	} else {
		gXml.selectSingleNode("/root/done").text := A_now								; Set <done> value to now
	}
	oExcel := oWorkbook.Application														; close workbook
	oExcel.DisplayAlerts := false
	oExcel.quit
	Return
}

	
;#endregion

;#region == PATIENT DIRECTORY HANDLING =====================================================================
patdir(LV,rownum) {
/* This is just a shell
*/	
	MsgBox
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

ParseDate(x) {
	mo := ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
	moStr := "Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec"
	dSep := "[ \-\._/]"
	date := {yyyy:"",mmm:"",mm:"",dd:"",date:""}
	time := {hr:"",min:"",sec:"",days:"",ampm:"",time:""}

	x := RegExReplace(x,"[,\(\)]")

	if (x~="\d{4}.\d{2}.\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?Z") {
		x := RegExReplace(x,"[TZ]","|")
	}
	if (x~="\d{4}.\d{2}.\d{2}T\d{2,}") {
		x := RegExReplace(x,"T","|")
	}

	if RegExMatch(x,"i)(\d{1,2})" dSep "(" moStr ")" dSep "(\d{4}|\d{2})",&d) {			; 03-Jan-2015
		date.dd := zdigit(d[1])
		date.mmm := d[2]
		date.mm := zdigit(objhasvalue(mo,d[2]))
		date.yyyy := d[3]
		date.date := trim(d[0])
	}
	else if RegExMatch(x,"\b(\d{4})[\-\.](\d{2})[\-\.](\d{2})\b",&d) {					; 2015-01-03
		date.yyyy := d[1]
		date.mm := zdigit(d[2])
		date.mmm := mo[d[2]]
		date.dd := zdigit(d[3])
		date.date := trim(d[0])
	}
	else if RegExMatch(x,"i)(" moStr "|\d{1,2})" dSep "(\d{1,2})" dSep "(\d{4}|\d{2})",&d) {	; Jan-03-2015, 01-03-2015
		date.dd := zdigit(d[2])
		date.mmm := objhasvalue(mo,d[1]) 
			? d[1]
			: mo[d[1]]
		date.mm := objhasvalue(mo,d[1])
			? zdigit(objhasvalue(mo,d[1]))
			: zdigit(d[1])
		date.yyyy := (d[3]~="\d{4}")
			? d[3]
			: (d[3]>50)
				? "19" d[3]
				: "20" d[3]
		date.date := trim(d[0])
	}
	else if RegExMatch(x,"i)(" moStr ")\s+(\d{1,2}),?\s+(\d{4})",&d) {					; Dec 21, 2018
		date.mmm := d[1]
		date.mm := zdigit(objhasvalue(mo,d[1]))
		date.dd := zdigit(d[2])
		date.yyyy := d[3]
		date.date := trim(d[0])
	}
	else if RegExMatch(x,"\b(19\d{2}|20\d{2})(\d{2})(\d{2})((\d{2})(\d{2})(\d{2})?)?\b",&d)  {	; 20150103174307 or 20150103
		date.yyyy := d[1]
		date.mm := d[2]
		date.mmm := mo[d[2]]
		date.dd := d[3]
		if (d[1]) {
			date.date := d[1] "-" d[2] "-" d[3]
		}
		
		time.hr := d[5]
		time.min := d[6]
		time.sec := d[7]
		if (d[5]) {
			time.time := d[5] ":" d[6] . strQ(d[7],":###")
		}
	}

	if RegExMatch(x,"i)(\d+):(\d{2})(:\d{2})?(:\d{2})?(.*)?(AM|PM)?",&t) {				; 17:42 PM
		hasDays := (t[4]) ? true : false 											; 4 nums has days
		time.days := (hasDays) ? t[1] : ""
		time.hr := trim(t[1+hasDays])
		time.min := trim(t[2+hasDays]," :")
		time.sec := trim(t[3+hasDays]," :")
		if (time.min>59) {
			time.hr := floor(time.min/60)
			time.min := zDigit(Mod(time.min,60))
		}
		if (time.hr>23) {
			time.days := floor(time.hr/24)
			time.hr := zDigit(Mod(time.hr,24))
			DHM:=true
		}
		time.ampm := trim(t[5])
		time.time := trim(t[0])
	}

	return {yyyy:date.yyyy, mm:date.mm, mmm:date.mmm, dd:date.dd, date:date.date
			, YMD:date.yyyy date.mm date.dd
			, YMDHMS:date.yyyy date.mm date.dd zDigit(time.hr) zDigit(time.min) zDigit(time.sec)
			, MDY:date.mm "/" date.dd "/" date.yyyy
			, MMDD:date.mm "/" date.dd 
			, hrmin:zdigit(time.hr) ":" zdigit(time.min)
			, days:zdigit(time.days)
			, hr:zdigit(time.hr), min:zdigit(time.min), sec:zdigit(time.sec)
			, ampm:time.ampm, time:time.time
			, DHM:zdigit(time.days) ":" zdigit(time.hr) ":" zdigit(time.min) " (DD:HH:MM)" 
			, DT:date.mm "/" date.dd "/" date.yyyy " at " zdigit(time.hr) ":" zdigit(time.min) ":" zdigit(time.sec) }
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
#Include xml.ahk
#Include strx2.ahk
; #Include MsgBox2.ahk
; #Include Class_LV_Colors.ahk
; #Include sift3.ahk
; #Include CMsgBox.ahk