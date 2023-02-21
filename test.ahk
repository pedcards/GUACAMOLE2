#Requires AutoHotkey v2

x := "This is a string."

y := xStr(x,,"","str")
z := stRegX(x,"this",1,1,"string",1)
q := StrX(x,"",0,0,"string",1,1)

MsgBox "'" y "'`n'" z "'`n'" q "'"

ExitApp

#Include "includes\strx2.ahk"
