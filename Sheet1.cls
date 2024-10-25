VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim fCalendarVisible As Boolean

Sub dateFormat()

    Range("A1:Z99").NumberFormat = "dd/mm/yyyy"
    
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    If fCalendarVisible Then Unload UserForm1

    If VBA.IsDate(Target.Cells(1, 1).Value) Or VBA.IsDate(Target.Cells(1, 1).Offset(-1)) Or InStr(1, UCase$(Target.Cells(1, 1).Offset(-1)), "DATE", vbTextCompare) > 0 Then
    
        ActiveSheet.Shapes("CalendarIcon").Top = Target.Top
        ActiveSheet.Shapes("CalendarIcon").Left = Target.Left + Target.Width + 8
        ActiveSheet.Shapes("CalendarIcon").Visible = True
        
    Else
    
        ActiveSheet.Shapes("CalendarIcon").Visible = False
        
    
    End If


End Sub

Sub showCal()

    UserForm1.Show False
    MoveFormToMouse
    buildCalendar
    fCalendarVisible = True

End Sub

Sub buildCalendar()

    Dim iYear As Integer
    Dim iMonth As Integer
    
    Dim startOfMonth As Date
    Dim trackingDate As Date
    
    Dim iStartOfMonthDay As Integer
    Dim cDay As Control
    
    If VBA.IsDate(ActiveCell.Value) Then
        
        trackingDate = ActiveCell.Value
        
    Else
    
        trackingDate = Now()
        
    End If
    
    iYear = VBA.Year(trackingDate)
    iMonth = VBA.Month(trackingDate)
    
    UserForm1.Controls("lblMonth").Caption = VBA.MonthName(iMonth, False)
    UserForm1.Controls("lblYear").Caption = iYear
    
    startOfMonth = VBA.DateSerial(iYear, iMonth, 1)
    iStartOfMonthDay = VBA.Weekday(startOfMonth, vbSunday)
    
    trackingDate = DateAdd("d", -iStartOfMonthDay + 1, startOfMonth)
    
    For i = 1 To 42
    
        Set cDay = UserForm1.Controls("day" & i)
        cDay.Caption = VBA.Day(trackingDate)
        cDay.Tag = trackingDate
        
        If VBA.Month(trackingDate) <> iMonth Then cDay.ForeColor = 8421504
        
        trackingDate = VBA.DateAdd("d", 1, trackingDate)
        
    Next i
    

End Sub

