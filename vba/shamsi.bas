' This class is the VBA version of alireza-ahmadi's javascript jalali conversion date file
' It Has 3 properties: y : year, m: month, d: day all jalali and  GDate : Date that in set
' it converts Georgian date to jalali and in get converts jalali date and returns Georgian
' date.
' Authr: M.samadi
' Date : 1393-10-25
' Version: 1.0

Private pYear As Integer
Private pMonth As Integer
Private pDay As Integer




Private Function isLeap() As Boolean
Dim m As Integer
m = pYear Mod 33
isLeap = (m = 1 Or m = 5 Or m = 9 Or m = 13 Or m = 17 Or m = 22 Or m = 26 Or m = 30)
End Function

Public Property Get y() As Integer
y = pYear
End Property
Public Property Let y(Value As Integer)
pYear = Value
End Property

Public Property Get m() As Integer
m = pMonth
End Property

Public Property Let m(Value As Integer)
pMonth = Value
End Property
Public Property Get d() As Integer
d = pDay
End Property
Public Property Let d(Value As Integer)
pDay = Value
End Property



Public Property Get GDate() As Date
    g_days_in_month = Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
    j_days_in_month = Array(31, 31, 31, 31, 31, 31, 30, 30, 30, 30, 30, 29)
    jy = pYear - 979
    jm = pMonth - 1
    jd = pDay - 1
    j_day_no = 365 * jy + Int(jy / 33) * 8 + Int(((jy Mod 33) + 3) / 4)
    
    For I = 0 To jm - 1
        j_day_no = j_day_no + j_days_in_month(I)
    Next I
    
    j_day_no = j_day_no + jd
    g_day_no = j_day_no + 79
    gy = 1600 + 400 * Int(g_day_no / 146097)
    g_day_no = g_day_no Mod 146097
    
    leap = True
    If g_day_no >= 36525 Then
        g_day_no = g_day_no - 1
        gy = gy + 100 * Int(g_day_no / 36524)
        g_day_no = g_day_no Mod 36524
        
        If g_day_no >= 365 Then
            g_day_no = g_day_no + 1
        Else
            leap = False
        End If
        
    End If
    
    gy = gy + 4 * Int(g_day_no / 1461)
    g_day_no = (g_day_no Mod 1461)
    
    If g_day_no >= 366 Then
        leap = False
        
        g_day_no = g_day_no - 1
        gy = gy + Int(g_day_no / 365)
        g_day_no = g_day_no Mod 365
    End If
    I = 0
    off = 0
    Do While g_day_no >= (g_days_in_month(I) + off)
        g_day_no = g_day_no - (g_days_in_month(I) + off)
        I = I + 1
        If I = 1 And leap Then
            off = 1
        Else
            off = 0
        End If
    Loop
    gm = I + 1
    gd = g_day_no + 1
    
    GDate = DateSerial(gy, gm, gd)
    
End Property

Public Property Let GDate(Value As Date)
    g_days_in_month = Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
    j_days_in_month = Array(31, 31, 31, 31, 31, 31, 30, 30, 30, 30, 30, 29)
    gy = Year(Value) - 1600
    gm = Month(Value) - 1
    gd = Day(Value) - 1
    
    g_day_no = 365 * gy + Int((gy + 3) / 4) - Int((gy + 99) / 100) + Int((gy + 399) / 400)
    
    For I = 0 To gm - 1
        g_day_no = g_day_no + g_days_in_month(I)
    Next I
    If (gm > 1 And (((gy Mod 4) = 0 And (gy Mod 100) <> 0) Or ((gy Mod 400) = 0))) Then
        g_day_no = g_day_no + 1
    End If
    g_day_no = g_day_no + gd
    
    j_day_no = g_day_no - 79
    
    j_np = Int(j_day_no / 12053)
    j_day_no = (j_day_no Mod 12053)
    
    jy = 979 + 33 * j_np + 4 * Int(j_day_no / 1461)
    
    j_day_no = (j_day_no Mod 1461)
    
    If (j_day_no >= 366) Then
        jy = jy + Int((j_day_no - 1) / 365)
        j_day_no = (j_day_no - 1) Mod 365
    End If
    I = 0
    While I < 11 And j_day_no >= j_days_in_month(I)
        j_day_no = j_day_no - j_days_in_month(I)
        I = I + 1
    Wend
    jm = I + 1
    jd = j_day_no + 1
    
    pYear = jy
    pMonth = jm
    pDay = jd
    
End Property




