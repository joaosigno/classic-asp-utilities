<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Public Function Format(ByVal expr, ByVal sFormat)
    If Left(sFormat, 1) = "<" Then
        expr = LCase(expr)
        sFormat = Mid(sFormat, 2)
    End If
    If Left(sFormat, 1) = ">" Then
        expr = UCase(expr)
        sFormat = Mid(sFormat, 2)
    End If
    If LCase(sFormat) = "long time" Then
        Format = FormatDateTime(expr, vbLongTime)
        Exit Function
    ElseIf LCase(sFormat) = "long date" Then
        Format = FormatDateTime(expr, vbLongDate)
        Exit Function
    End If
    If InStr(sFormat, "0") + InStr(sFormat, "#") > 0 Then
        sFormat = vbsFormatNumber(expr, sFormat)
    End If
    Format = vbsFormatDate(expr, sFormat)
End Function

Private Function vbsFormatDate(ByVal expr, ByVal sFormat)
    Dim i, s, skip, tmp, prevhour, military
    i = 1
    prevhour = 0
    military = InStr ( LCase(sFormat), "am/pm" ) = 0 And InStr ( LCase(sFormat), "a/p" ) = 0
    While i <= Len(sFormat)
        skip = 0
        Select Case LCase(Mid(sFormat, i, 1))
        Case "a"
            If LCase(Mid(sFormat, i, 3)) = "a/p" Then
                If Hour(expr) < 12 Then
                    If Mid(sFormat, i, 1) = "A" Then
                        s = s & "A"
                    Else
                        s = s & "a"
                    End If
                Else
                    If Mid(sFormat, i, 1) = "A" Then
                        s = s & "P"
                    Else
                        s = s & "p"
                    End If
                End If
                skip = 3
                prevhour = 0
            ElseIf LCase(Mid(sFormat, i, 5)) = "am/pm" Then
                If Hour(expr) < 12 Then
                    If Mid(sFormat, i, 1) = "A" Then
                        s = s & "AM"
                    Else
                        s = s & "am"
                    End If
                Else
                    If Mid(sFormat, i, 1) = "A" Then
                        s = s & "PM"
                    Else
                        s = s & "pm"
                    End If
                End If
                skip = 5
                prevhour = 0
            End If
        Case "c"
            s = s & CStr(CDate(expr))
            skip = 1
            prevhour = 0
        Case "d"
            If LCase(Mid(sFormat, i, 3)) = "ddd" Then
                If LCase(Mid(sFormat, i+3, 1)) = "d" Then
                    skip = 4
                Else
                    skip = 3
                End If
                Select Case Mid(sFormat,i,2)
                Case "DD"
                    s = s & UCase(WeekdayName(Weekday(expr), skip<>4))
                Case "dd"
                    s = s & LCase(WeekdayName(Weekday(expr), skip<>4))
                Case Else
                    s = s & WeekdayName(Weekday(expr), skip<>4)
                End Select
            Else
                tmp = Day(expr)
                If LCase(Mid(sFormat, i + 1, 1)) = "d" Then
                    If 1 = Len(tmp) Then s = s & "0"
                    skip = 2
                Else
                    skip = 1
                End If
                s = s & tmp
            End If
            prevhour = 0
        Case "h"
            tmp = Hour(expr)
            If Not military Then tmp = ((tmp+11) Mod 12) + 1
            If LCase(Mid(sFormat, i + 1, 1)) = "h" Then
                If 1 = Len(tmp) Then s = s & "0"
                skip = 2
            Else
                skip = 1
            End If
            s = s & tmp
            prevhour = 1
        Case "m"
            If LCase(Mid(sFormat, i, 3)) = "mmm" Then
                If LCase(Mid(sFormat, i+3, 1)) = "m" Then
                    skip = 4
                Else
                    skip = 3
                End If
                Select Case Mid(sFormat,i,2)
                Case "MM"
                    s = s & UCase(MonthName(Month(expr), skip<>4))
                Case "mm"
                    s = s & LCase(MonthName(Month(expr), skip<>4))
                Case Else
                    s = s & MonthName(Month(expr), skip<>4)
                End Select
            Else
                If prevhour = 0 Then
                    tmp = Month(expr)
                Else
                    tmp = Minute(expr)
                End If
                If LCase(Mid(sFormat, i + 1, 1)) = "m" Then
                    If 1 = Len(tmp) Then s = s & "0"
                    skip = 2
                Else
                    skip = 1
                End If
                s = s & tmp
            End If
            prevhour = 0
        Case "n"
            tmp = Minute(expr)
            If LCase(Mid(sFormat, i + 1, 1)) = "n" Then
                If 1 = Len(tmp) Then s = s & "0"
                skip = 2
            Else
                skip = 1
            End If
            s = s & tmp
            prevhour = 0
        Case "q" ' quarter
            s = s & DatePart("q", expr)
            skip = 1
            prevhour = 0
        Case "s"
            tmp = Second(expr)
            If LCase(Mid(sFormat, i + 1, 1)) = "s" Then
                If 1 = Len(tmp) Then s = s & "0"
                skip = 2
            Else
                skip = 1
            End If
            s = s & tmp
            prevhour = 0
        Case "w" ' weekday
            If LCase(Mid(sFormat, i, 2)) = "ww" Then
                s = s & DatePart("ww", expr) ' week of year
                skip = 2
            Else
                s = s & Weekday(CDate(expr)) ' day of week
                skip = 1
            End If
            prevhour = 0
        Case "y"
            If LCase(Mid(sFormat, i, 4)) = "yyyy" Then
                s = s & Year(expr)
                skip = 4
            ElseIf LCase(Mid(sFormat, i, 2)) = "yy" Then
                s = s & Right(Year(expr), 2)
                skip = 2
            ElseIf LCase(Mid(sFormat, i, 1)) = "y" Then
                ' day # of the year
                s = s & DatePart("y", expr)
                skip = 1
            End If
            prevhour = 0
        End Select
        If 0 = skip Then
            s = s & Mid(sFormat, i, 1)
            skip = 1
        End If
        i = i + skip
    Wend
    vbsFormatDate = s
End Function

Function vbsFormatNumber(expr, sFormat)
    Dim src, fmt, comma, tmp, dst
    fmt = Split(CStr(sFormat), ".")
    ReDim Preserve fmt(1)
    expr = Round(expr, Len(fmt(1)))
    src = Split(CStr(expr), ".")
    ReDim Preserve src(1)
    ReDim dst(1)
    comma = InStr(fmt(0), ",") > 0
    If comma Then fmt(0) = Replace(fmt(0), ",", "")
    Do
        tmp = Replace(fmt(0), "0#", "00")
        If tmp = fmt(0) Then Exit Do
        fmt(0) = tmp
    Loop While True
    tmp = ""
    
    ' reverse first half and parse it...
    fmt(0) = StrReverse(fmt(0))
    src(0) = StrReverse(src(0))
    If src(0) = "0" Then src(0) = ""
    While src(0) <> ""
        tmp = Left(fmt(0), 1)
        If InStr("0#", tmp) Then
            dst(0) = dst(0) & Left(src(0), 1)
            src(0) = Mid(src(0), 2)
        Else
            dst(0) = dst(0) & tmp
        End If
        fmt(0) = Mid(fmt(0), 2)
    Wend
    While fmt(0) <> ""
        tmp = Left(fmt(0), 1)
        If tmp = "0" Then
            dst(0) = dst(0) & "0"
        ElseIf tmp <> "#" Then
            dst(0) = dst(0) & tmp
        End If
        fmt(0) = Mid(fmt(0), 2)
    Wend
    ' process commas if necessary
    If comma Then
        tmp = dst(0)
        dst(0) = ""
        While Len(tmp) > 3
            dst(0) = dst(0) & Left(tmp, 3) & ","
            tmp = Mid(tmp, 4)
        Wend
        dst(0) = dst(0) & tmp
        tmp = ""
    End If
    
    ' process second half
    While src(1) <> ""
        tmp = Left(fmt(0), 1)
        If InStr("0#", tmp) Then
            dst(1) = dst(1) & Left(src(1), 1)
            src(1) = Mid(src(1), 2)
        Else
            dst(1) = dst(1) & tmp
        End If
        fmt(1) = Mid(fmt(1), 2)
    Wend
    While fmt(1) <> ""
        tmp = Left(fmt(1), 1)
        If tmp = "0" Then
            dst(1) = dst(1) & "0"
        ElseIf tmp <> "#" Then
            dst(1) = dst(1) & tmp
        End If
        fmt(1) = Mid(fmt(1), 2)
    Wend

    tmp = StrReverse(dst(0))
    If dst(1) <> "" Then
        tmp = tmp & "." & dst(1)
    End If
    vbsFormatNumber = tmp
End Function
%>
