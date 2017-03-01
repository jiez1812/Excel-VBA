Attribute VB_Name = "Module1"
Sub button_click()
    With Sheets("Calendar").Range("B2:M38")
        .Interior.ColorIndex = 0
        .ClearContents
    End With
    create_cal
End Sub

Sub create_cal()
    Dim month_day(1 To 12) As Variant
    Dim r, d As Integer
    Dim year As Integer
    Dim get_weekday, day_week As String
    
    year = Sheets("Calendar").Cells(1, 1)
    get_weekday = find_weekday(year)
    
    month_day(1) = 31
    If (year Mod 4) = 0 Then
        month_day(2) = 29
    Else
        month_day(2) = 28
    End If
    month_day(3) = 31
    month_day(4) = 30
    month_day(5) = 31
    month_day(6) = 30
    month_day(7) = 31
    month_day(8) = 31
    month_day(9) = 30
    month_day(10) = 31
    month_day(11) = 30
    month_day(12) = 31
    
    For w = 2 To 8
        If Sheets("Calendar").Cells(w, 1).Value = get_weekday Then
            r = w
            Exit For
        End If
    Next w
    
    For i = 2 To 13
        d = 1
        If r > 2 Then
            For x = 2 To r - 1
                Sheets("Calendar").Cells(x, i).Interior.Color = RGB(216, 216, 216)
            Next x
        End If
        For j = r To (month_day(i - 1)) + (r - 1)
            Sheets("Calendar").Cells(j, i).Value = d
            Call check_weekend(CInt(j), CInt(i))
            d = d + 1
        Next j
        For x = j To 38
            Sheets("Calendar").Cells(x, i).Interior.Color = RGB(216, 216, 216)
        Next x
        r = (find_weekday2(Sheets("Calendar").Cells((month_day(i - 1)) + (r - 1), 1).Value)) + 1
        If r = 8 Then
            r = 1
        End If
        r = r + 1
    Next i
End Sub

Function find_weekday(yearnum As Integer) As String
    Dim j As Integer
    Dim w As Date
    
    w = "1-1-" + CStr(yearnum)
    j = Weekday(w)
    
    
    Select Case j
        Case 2
            find_weekday = "Mon"
        Case 3
            find_weekday = "Tue"
        Case 4
            find_weekday = "Wed"
        Case 5
            find_weekday = "Thu"
        Case 6
            find_weekday = "Fri"
        Case 7
            find_weekday = "Sat"
        Case 1
            find_weekday = "Sun"
    End Select
End Function

Function find_weekday2(DayWeek As String) As Integer
    Select Case DayWeek
        Case "Mon"
            find_weekday2 = 3
        Case "Tue"
            find_weekday2 = 4
        Case "Wed"
            find_weekday2 = 5
        Case "Thu"
            find_weekday2 = 6
        Case "Fri"
            find_weekday2 = 7
        Case "Sat"
            find_weekday2 = 1
        Case "Sun"
            find_weekday2 = 2
    End Select
End Function

Function check_weekend(rows As Integer, cols As Integer)
    If Sheets("Calendar").Cells(rows, 1).Value = "Sun" Or Sheets("Calendar").Cells(rows, 1).Value = "Sat" Then
        Sheets("Calendar").Cells(rows, cols).Interior.Color = RGB(252, 213, 180)
    End If
End Function
