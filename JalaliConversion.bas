' This file is a simple use of shamsi.bas to add two function of toShamsi and toMiladi
' to excel.
' toShamsi gets a date and optinal delimiter and returns a string jalali date in 
' "YYYY/MM/DD" format.
' toMiladi gets a jalali date string, an optional delimiter and a format string
' and returns a georgian date string based on specified date format.
' Auther: M.samadi
' Date : 1393-10-25
' Version: 1.0
' 
Function toShamsi(ID As Date, Optional delimiter As String = "/") As String
Dim s As Shamsi
Set s = New Shamsi
s.GDate = ID
toShamsi = s.y & delimiter & s.m & delimiter & s.d
End Function

Function toMiladi(sdate As String, Optional delimiter As String = "/", Optional sformat As String = "M") As String
Dim s As Shamsi
Set s = New Shamsi
sd = Split(sdate, delimiter, 3)
y = Val(sd(0))
If y < 1300 Then
    y = y + 1300
End If
m = Val(sd(1))
d = Val(sd(2))

s.y = y
s.m = m
s.d = d
toMiladi = s.GDate
End Function
