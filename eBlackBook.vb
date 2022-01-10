Option Explicit
Public globalpass As Boolean

Private Sub Workbook_Open()
    ''''Change globalpass from True to False to unprotect the ENTIRE workbook''''
    ''''Caution warranted''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    globalpass = True
    ''''After changing, save and restart notebook''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Application.ScreenUpdating = False
    Dim i
    For i = 1 To Worksheets.count
        Worksheets(i).Range("A1:B5").ClearComments
    Next
    Sheet13.Range("E2").value = DateTime.Date
    Sheet14.Range("F1").value = DateTime.Date
    Sheet14.Activate
    Call SetTimer
    Call FetchBB_Click
    Sheet13.Activate
    'MsgBox ("Welcome to electronicBB! Forgot to put this in my email so I'm putting it here. No one else can open this file while you have it open, so I have it set to save/close after 15 minutes of inactivity for now. fyi")
    Application.ScreenUpdating = True
End Sub

'see IdleSaveClose for scripts used below'
'

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call StopTimer
End Sub


Private Sub Workbook_SheetCalculate(ByVal Sh As Object)
    Call StopTimer
    Call SetTimer
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, _
  ByVal Target As Excel.Range)
    Call StopTimer
    Call SetTimer
End Sub

Sub Button1_Click()
    Dim selectdate
    Dim noteForm As New UserFormNote
    selectdate = Application.ActiveCell
    If ActiveCell.Column = 1 Or ActiveCell.Column > 8 Then
        MsgBox ("Please select the date first")
        Exit Sub
    ElseIf ActiveCell.Row = 3 Or ActiveCell.Row = 5 Or ActiveCell.Row = 7 Or ActiveCell.Row = 9 Or ActiveCell.Row = 11 Or ActiveCell.Row = 13 Then
        'user selected date form
        noteForm.textDate = ActiveCell
        noteForm.Show
    ElseIf ActiveCell.Row = 4 Or ActiveCell.Row = 6 Or ActiveCell.Row = 8 Or ActiveCell.Row = 10 Or ActiveCell.Row = 12 Or ActiveCell.Row = 14 Then
        'user select date descript'
        noteForm.textDate = ActiveCell.Offset(-1)
        noteForm.Show
    Else
        MsgBox ("Please select the date first")
    End If
End Sub

Public Sub AddTask(ByVal taskie As String)
    'Day 3 Read Chr(128) 3/31/2021 Chr(138) notes_optional'
    Dim hand, notes
    task = Split(taskie, Chr(128))(0)
    notes = Split(Split(taskie, Chr(128))(1), Chr(138))(1)
    If notes <> "" Then
        task = task & " <" & notes & ">"
    End If
    Dim i As Integer
    For i = 1 To Worksheets.count
        If Worksheets(i).CodeName <> "Sheet13" And Worksheets(i).CodeName <> "Sheet14" Then
            For Each c In Worksheets(i).Range("B3:H3,B5:H5,B7:H7,B9:H9,B11:H11,B13:C13")
                If FormatDateTime(c) = Split(Split(taskie, Chr(128))(1), Chr(138))(0) Then
                    'see if cell is empty, if so add first task without space'
                    If Worksheets(i).Cells(c.Row + 1, c.Column) = "" Then
                        Worksheets(i).Unprotect
                        Worksheets(i).Cells(c.Row + 1, c.Column) = task
                        Worksheets(i).Protect
                    Else
                        Worksheets(i).Unprotect
                        Worksheets(i).Cells(c.Row + 1, c.Column) = Worksheets(i).Cells(c.Row + 1, c.Column).value & Chr(10) & task
                        Worksheets(i).Protect
                    End If
                End If
            Next
         End If
    Next
    'Sheet1.Range("B3:H3,B5:H5,B7:H7,B9:H9,B11:H11,B13:C13")'
End Sub

Function getBatch(ByVal usedate As Date) As String
    Dim thingy() As String
    Dim Day, month, year As String
    'ex 4/17/2021'
    thingy = Split(usedate, "/")
    month = thingy(0)
    If Len(month) = 1 Then
        month = "0" & month
    End If
    Day = thingy(1)
    If Len(Day) = 1 Then
        Day = "0" & Day
    End If
    year = Mid(thingy(2), 3)
    'returns 041721'
    getBatch = month & Day & year
End Function

Function getColor(ByVal weekdy As String)
    'get color for weekday'
    Select Case weekdy
        Case "MON"
            'blue'
            getColor = "&HFFFF00"
        Case "TUE"
            'orange'
            getColor = &H80FF&
        Case "WED"
            'yellow'
            getColor = &HFFFF&
        Case "THU"
            'green - source sarah'
            getColor = &HFF00&
        Case "FRI"
            'Purple'
            getColor = &HFF00FF
        Case "SAT"
            'red'
            getColor = &HFF&
        Case "SUN"
            'white'
            getColor = &H8000000E
        Case "ERR"
            'pink for error'
            getColor = &HFFC0FF
    End Select
End Function

Function GetDay(ByVal usedate As Date, ByVal DayX As Integer)
    Dim Day, month, year As String
    'ex 4/17/2021'
    Day = Split(usedate, "/")(1)
    month = Split(usedate, "/")(0)
    year = Split(usedate, "/")(2)
    If Sheet1.Range("K3") <> year Then
        MsgBox ("Wrong year! Wrong Calendar")
        Return
    End If
    'make sure year is correct'
    Dim datex As Date
    datex = usedate + DayX
    Dim weekcolor, weekday, todo As String
    Dim i As Integer
    Dim makeshort(2)
    For i = 1 To Worksheets.count
        If Worksheets(i).CodeName <> "Sheet13" And Worksheets(i).CodeName <> "Sheet14" Then
            'Sheet1.Range("B3:H3,B5:H5,B7:H7,B9:H9,B11:H11,B13:C13")'
            For Each c In Worksheets(i).Range("B3:H3,B5:H5,B7:H7,B9:H9,B11:H11,B13:C13")
                If c = datex Then
                    'MsgBox ("for dayx =3, got the third day")'
                    weekcolor = Left(Worksheets(i).Cells(2, c.Column), 3)
                    weekday = Replace(c, "/" & year, "")
                    todo = Worksheets(i).Cells(c.Row + 1, c.Column)
                    makeshort(0) = weekcolor
                    makeshort(1) = weekday
                    makeshort(2) = todo
                    GetDay = makeshort
                    Exit Function
                    'grab the revelant info'
                End If
            Next
        End If
    Next
    makeshort(0) = "ERR"
    makeshort(1) = "OR"
    makeshort(2) = ""
    GetDay = makeshort
End Function

Function ShortestTime(ByVal usedate As Date, ByVal firstday As Integer, ByVal lastday As Integer) As Integer
    Dim shorttime As Integer
    Dim d As Integer
    Dim worktime As Integer
    Dim shortd As Integer
    Dim value
    'value holds getDay results'
    'value(0) = weekcolor'
    'value(1) = weekday'
    'value(2) = todo'
    shorttime = 10000
    For d = firstday To lastday
        worktime = 0
        value = GetDay(usedate, d)
        If InStr(value(2), Chr(10)) = 0 And value(2) = "" Then
            'do nothing vvv old code for manual entry that lacks chr(10)'
            'worktime = Val(Replace((Split(value(2), "(")(1)), ")", ""))'
            'mfForm(Replace("listD" & Str(d), " ", "")).AddItem (Replace((Split(value(2), "(")(0)), ")", ""))'
        Else
            work = Split(value(2), Chr(10))
            For Each wut In work
                'MsgBox (wut)
                If wut = "" Or InStr(wut, "(") = 0 Then
                    'do nothing'
                Else
                    worktime = worktime + Val(Replace((Split(wut, "(")(1)), ")", ""))
                End If
                'worktime = worktime + getTime()'
            Next
        End If
        If worktime < shorttime Then
            shortd = d
            shorttime = worktime
        ElseIf worktime = shorte And value(0) <> "SAT" And value(0) <> "SUN" Then
            'select for non weekend day if possible'
            shortd = d
            shorttime = worktime
        End If
    Next
    ShortestTime = shortd
End Function

Function getTime(ByVal count As String, ByVal typ As String) As String
    Dim num, result As Long
    If IsNumeric(count) = False Or Len(count) = 0 Then
        'check if number of sample/baskets is actually a number'
        getTime = "0"
        Exit Function
    Else
        num = FormatNumber(count)
        ' insert math here later!!!!! '
        'mf - below 8 canisters 1 minute, after that y = 0.1591x2 - 1.9591x + 6.5182 where y is time and x is #'
        'dt/nonIso - y = 1.8647e0.183x
        'move em manually set to 2
        'out em manually set to 5
        Select Case typ
            Case "mf"
                If num <= 4 Then
                    result = 2
                ElseIf num <= 10 Then
                    result = 3
                ElseIf num <= 20 Then
                    result = 4
                ElseIf num > 20 Then
                    result = 5
                End If
            Case "dt"
                If num <= 2 Then
                    result = 4
                ElseIf num <= 6 Then
                    result = 8
                ElseIf num <= 10 Then
                    result = 10
                ElseIf num <= 15 Then
                    result = 16
                ElseIf num > 15 Then
                    result = 30
                End If
            Case "dtt"
                If num <= 2 Then
                    result = 4
                ElseIf num <= 6 Then
                    result = 8
                ElseIf num <= 10 Then
                    result = 10
                ElseIf num <= 15 Then
                    result = 16
                ElseIf num > 15 Then
                    result = 30
                End If
            Case "noniso"
                result = Round(num, 0) * 5
            Case Else
                result = Round(num, 0)
        End Select
    End If
    getTime = Replace(Str(result), " ", "")
End Function

Dim DownTime As Date

Sub SetTimer()
    DownTime = Now + TimeValue("00:15:00")
    Application.OnTime EarliestTime:=DownTime, _
      Procedure:="ShutDown", Schedule:=True
End Sub

Sub StopTimer()
    On Error Resume Next
    Application.OnTime EarliestTime:=DownTime, _
      Procedure:="ShutDown", Schedule:=False
 End Sub

Sub ShutDown()
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    With ThisWorkbook
        .Saved = True
        .Close
    End With
End Sub

'see ThisWorkbook obj for triggers'

Sub Button19_Click()
    Dim i As Integer
    For i = 3 To 90
        If Sheet14.Cells(i, 3) = "Pending" Then
            Sheet14.Cells(i, 3) = "Complete"
        End If
    Next
    Dim temp
    temp = Sheet14.Range("C3")
    Sheet14.Range("C3") = ""
    Sheet14.Range("C3") = temp
End Sub

Sub FetchBB_Click()
    Dim usedate As Date
    usedate = Sheet14.Range("F1").value
    Sheet14.Cells(20, 40) = "working"
    Application.ScreenUpdating = False
    Range("A1:D1").Select
    'clear previous table and re-format
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Sheet14.Range("A2:D90").Select
    Selection.Clear
    Sheet14.Range("A2:D2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Sheet14.Range("A3:A90").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .Rows.AutoFit
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Sheet14.Range("B3:B90").Select
    With Selection
        .Rows.AutoFit
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Sheet14.Range("C3:C90").Select
    With Selection
        .Rows.AutoFit
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Sheet14.Range("D3:D90").Select
    With Selection
        .Rows.AutoFit
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .RowHeight = 20
    End With
    Sheet14.Cells(2, 1) = "Task"
    Sheet14.Cells(2, 2) = "Time"
    Sheet14.Cells(2, 3) = "Status"
    Sheet14.Cells(2, 4) = "Notes"
    Dim r As Integer
    Dim tasks
    tasks = Split(getTasks(usedate), Chr(10))
    Dim donelist, stat As String
    donelist = getTaskStatus(usedate, tasks)
    Sheet14.Cells(51, 50) = usedate
    For Each task In tasks
        'MF042921 Day 3 Read'
        '(69)'
        stat = Mid(donelist, r + 1, 1)
        If stat = "" Then
            stat = "P"
        End If
        Sheet14.Cells(3 + r, 50) = stat
        If InStr(task, "(") <> 0 And InStr(task, "<") <> 0 Then
            'MF042921 Day 3 Read (69) <something>'
            Sheet14.Cells(3 + r, 1) = Split(task, "(")(0)
            Sheet14.Cells(3 + r, 2) = Split(Split(task, "(")(1), ")")(0)
            Sheet14.Cells(3 + r, 4) = Split(Split(task, "<")(1), ">")
            Call AddDropDown(3 + r, 3, stat)
        ElseIf InStr(task, "(") = 0 And InStr(task, "<") <> 0 Then
            'PTO BH <in Canada eh!>'
            Sheet14.Cells(3 + r, 1) = Split(task, "<")(0)
            Sheet14.Cells(3 + r, 4) = Split(Split(task, "<")(1), ">")
            Sheet14.Cells(3 + r, 3) = "Communication"
            Sheet14.Cells(3 + r, 1).Style = "Neutral"
            Sheet14.Cells(3 + r, 2).Style = "Neutral"
            Sheet14.Cells(3 + r, 3).Style = "Neutral"
            Sheet14.Cells(3 + r, 4).Style = "Neutral"
        ElseIf InStr(task, "(") = 0 And InStr(task, "<") = 0 Then
            'PTO BH'
            Sheet14.Cells(3 + r, 1) = task
            Sheet14.Cells(3 + r, 3) = "Communication"
            Sheet14.Cells(3 + r, 1).Style = "Neutral"
            Sheet14.Cells(3 + r, 2).Style = "Neutral"
            Sheet14.Cells(3 + r, 3).Style = "Neutral"
            Sheet14.Cells(3 + r, 4).Style = "Neutral"
        ElseIf InStr(task, "(") <> 0 And InStr(task, "<") = 0 Then
            'MF042921 Day 3 Read (1)'
            Sheet14.Cells(3 + r, 1) = Split(task, "(")(0)
            Sheet14.Cells(3 + r, 2) = Split(Split(task, "(")(1), ")")(0)
            Call AddDropDown(3 + r, 3, stat)
        Else
            MsgBox ("Error: Imported task did not conform to format expectation" & Chr(10) & task & Chr(10) & "Please re-format entry in calendar and refresh")
            Exit Sub
        End If
        r = r + 1
    Next
    Sheet14.ListObjects.Add(xlSrcRange, Range("$A$2:$D$" & Replace(CStr(2 + r), " ", "")), , xlYes).Name = _
        "Table1"
    Sheet14.ListObjects("Table1").TableStyle = "TableStyleMedium14"
    Sheet14.Range("A3:D90").Select
    With Selection
        .Rows.AutoFit
    End With
    Sheet14.Range("A3:D90").Select
    With Selection
        .RowHeight = .RowHeight * 1.5
    End With
    Sheet14.Cells(20, 40) = ""
    Range("F1").Select
    If getDayStatus(usedate) Then
        Call Button19_Click
    End If
    Application.ScreenUpdating = True
End Sub

Function getTasks(ByVal usedate As String)
    Dim i As Integer
    For i = 1 To Worksheets.count
        If Worksheets(i).CodeName <> "Sheet13" And Worksheets(i).CodeName <> "Sheet14" Then
            For Each c In Worksheets(i).Range("B3:H3,B5:H5,B7:H7,B9:H9,B11:H11,B13:C13")
                If FormatDateTime(c) = usedate Then
                    getTasks = Worksheets(i).Cells(c.Row + 1, c.Column)
                    Exit Function
                End If
            Next
         End If
    Next
End Function

'MF042921 Day 3 Read (1) <some ome>
'MF042921 Day 3 Read (1)

Sub AddDropDown(ByVal r As Integer, ByVal c As Integer, ByVal stat As String)
    Sheet14.Cells(r, c).Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="Pending,Complete"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    If stat = "P" Then
        Sheet14.Cells(r, c) = "Pending"
    Else
        Sheet14.Cells(r, c) = "Complete"
        Sheet14.Cells(r, 1).Style = "Good"
        Sheet14.Cells(r, 2).Style = "Good"
        Sheet14.Cells(r, 3).Style = "Good"
        Sheet14.Cells(r, 4).Style = "Good"
    End If
End Sub

Function getDayStatus(ByVal usedate As Date) As Boolean
    Dim i As Integer
    For i = 1 To Worksheets.count
        If Worksheets(i).CodeName <> "Sheet13" And Worksheets(i).CodeName <> "Sheet14" Then
            'Sheet1.Range("B3:H3,B5:H5,B7:H7,B9:H9,B11:H11,B13:C13")'
            For Each c In Worksheets(i).Range("B3:H3,B5:H5,B7:H7,B9:H9,B11:H11,B13:C13")
                If c = usedate Then
                    'MsgBox ("for dayx =3, got the third day")'
                    If Worksheets(i).Cells(c.Row + 1, c.Column).Style = "Good" Then
                        getDayStatus = True
                    Else
                        getDayStatus = False
                    End If
                    'grab the revelant info'
                End If
            Next
        End If
    Next
End Function

Function getTaskStatus(ByVal usedate As String, ByVal tasks As Variant) As String
    Dim i As Integer
    Dim temp As String
        For i = 1 To Worksheets.count
            If Worksheets(i).CodeName <> "Sheet13" And Worksheets(i).CodeName <> "Sheet14" Then
                'Sheet1.Range("B3:H3,B5:H5,B7:H7,B9:H9,B11:H11,B13:C13")'
                For Each c In Worksheets(i).Range("B3:H3,B5:H5,B7:H7,B9:H9,B11:H11,B13:C13")
                    If c = usedate Then
                        temp = Worksheets(i).Cells(c.Row + 50, c.Column)
                        If temp = "" Then
                            For Each task In tasks
                                temp = temp & "P"
                            Next
                            Worksheets(i).Unprotect
                            Worksheets(i).Cells(c.Row + 50, c.Column) = temp
                            Worksheets(i).Protect
                        End If
                        getTaskStatus = Worksheets(i).Cells(c.Row + 50, c.Column)
                    End If
                Next
            End If
        Next
End Function



Private Sub updatePend()
    Dim value, usedate, work, worktime
    Me.listpending.Clear
    Me.buttonBack.Visible = True
    Me.buttonForw.Visible = True
    usedate = Sheet13.Range("E2").value
    value = GetDay(usedate, Me.labelPending.Tag)
    'value holds the following:'
        'value(0) = weekcolor
        'value(1) = weekday'
        'value(2) = todo'
    If InStr(value(2), Chr(10)) = 0 And value(2) = "" Then
        Me.listpending.AddItem ("N/A")
        Me.timepending.Text = "Estimated Time: 0 mins"
        'do nothing vvv old code for manual entry that lacks chr(10)'
        'worktime = Val(Replace((Split(value(2), "(")(1)), ")", ""))'
        'mfForm(Replace("listD" & Str(d), " ", "")).AddItem (Replace((Split(value(2), "(")(0)), ")", ""))'
    Else
         'Day 3 Read Chr(128) 3/31/2021 Chr(138) notes_optional'
        work = Split(value(2), Chr(10))
        For Each wut In work
            Me.listpending.AddItem (wut)
            If wut = "" Or InStr(wut, "(") = 0 Then
                'do nothing'
            Else
                worktime = worktime + Val(Replace((Split(wut, "(")(1)), ")", ""))
            End If
            'worktime = worktime + getTime()'
        Next
        Me.timepending.Text = "Estimated Time: " & CStr(worktime) & " mins"
    End If
End Sub

'''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''
'''Navigating Buttons 4 Pending'''''
'''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''

Private Sub buttonBack_Click()
    Dim value, usedate
    usedate = Sheet13.Range("E2").value
    Me.labelPending.Tag = Me.labelPending.Tag - 1
    value = GetDay(usedate, Me.labelPending.Tag)
    Me.labelPending.Caption = value(0) & " " & value(1)
    Me.labelPending.BackColor = getColor(value(0))
    Call updatePend
End Sub

Private Sub buttonForw_Click()
    Dim value, usedate
    usedate = Sheet13.Range("E2").value
    Me.labelPending.Tag = Me.labelPending.Tag + 1
    value = GetDay(usedate, Me.labelPending.Tag)
    Me.labelPending.Caption = value(0) & " " & value(1)
    Me.labelPending.BackColor = getColor(value(0))
    Call updatePend
End Sub


'''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''
'''''''''Changing Day #  ''''''''''
'''''''''''''''''''''''''''''''''''

Private Sub day1_Change()
    Dim timee As Integer, usedate As Date
    Dim value
    usedate = Sheet13.Range("E2").value
    'value holds the following:'
        'value(0) = weekcolor'
        'value(1) = weekday'
        'value(2) = todo'
    If IsNumeric(day1.Text) And IsEmpty(day1.Text) = False Then
        Me.weekday1.Visible = True
        value = GetDay(usedate, CInt(Me.day1))
        Me.weekday1 = value(0) & " " & value(1)
        Me.weekday1.BackColor = getColor(value(0))
    ElseIf Me.day1 = "" And Me.Tag = "cust" Then
        Me.weekday1.Visible = False
        Me.task1 = ""
        Me.time1 = ""
    Else
        Me.day1 = ""
        Me.weekday1.Visible = False
    End If
    If Me.weekday1 = "ERR OR" Then
        Me.day1 = ""
        Me.weekday1.Visible = False
    End If
End Sub

Private Sub day2_Change()
    Dim timee As Integer, usedate As Date
    Dim value
    usedate = Sheet13.Range("E2").value
    'value holds the following:'
        'value(0) = weekcolor'
        'value(1) = weekday'
        'value(2) = todo'
    If IsNumeric(day2.Text) And IsEmpty(day2.Text) = False Then
        Me.weekday2.Visible = True
        value = GetDay(usedate, CInt(Me.day2))
        Me.weekday2 = value(0) & " " & value(1)
        Me.weekday2.BackColor = getColor(value(0))
    ElseIf Me.day2 = "" And Me.Tag = "cust" Then
        Me.weekday2.Visible = False
        Me.task2 = ""
        Me.time2 = ""
    Else
        Me.day2 = ""
        Me.weekday2.Visible = False
    End If
    If Me.weekday2 = "ERR OR" Then
        Me.day2 = ""
        Me.weekday2.Visible = False
    End If
End Sub

Private Sub day3_Change()
    Dim timee As Integer, usedate As Date
    Dim value
    usedate = Sheet13.Range("E2").value
    'value holds the following:'
        'value(0) = weekcolor'
        'value(1) = weekday'
        'value(3) = todo'
    If IsNumeric(day3.Text) And IsEmpty(day3.Text) = False Then
        Me.weekday3.Visible = True
        value = GetDay(usedate, CInt(Me.day3))
        Me.weekday3 = value(0) & " " & value(1)
        Me.weekday3.BackColor = getColor(value(0))
    ElseIf Me.day3 = "" And Me.Tag = "cust" Then
        Me.weekday3.Visible = False
        Me.task3 = ""
        Me.time3 = ""
    Else
        Me.day3 = ""
        Me.weekday3.Visible = False
    End If
    If Me.weekday3 = "ERR OR" Then
        Me.day3 = ""
        Me.weekday3.Visible = False
    End If
End Sub

Private Sub day4_Change()
    Dim timee As Integer, usedate As Date
    Dim value
    usedate = Sheet13.Range("E2").value
    'value holds the following:'
        'value(0) = weekcolor'
        'value(1) = weekday'
        'value(4) = todo'
    If IsNumeric(day4.Text) And IsEmpty(day4.Text) = False Then
        Me.weekday4.Visible = True
        value = GetDay(usedate, CInt(Me.day4))
        Me.weekday4 = value(0) & " " & value(1)
        Me.weekday4.BackColor = getColor(value(0))
    ElseIf Me.day4 = "" And Me.Tag = "cust" Then
        Me.weekday4.Visible = False
        Me.task4 = ""
        Me.time4 = ""
    Else
        Me.day4 = ""
        Me.weekday4.Visible = False
    End If
    If Me.weekday4 = "ERR OR" Then
        Me.day4 = ""
        Me.weekday4.Visible = False
    End If
End Sub

Private Sub day5_Change()
    Dim timee As Integer, usedate As Date
    Dim value
    usedate = Sheet13.Range("E2").value
    'value holds the following:'
        'value(0) = weekcolor'
        'value(1) = weekday'
        'value(5) = todo'
    If IsNumeric(day5.Text) And IsEmpty(day5.Text) = False Then
        Me.weekday5.Visible = True
        value = GetDay(usedate, CInt(Me.day5))
        Me.weekday5 = value(0) & " " & value(1)
        Me.weekday5.BackColor = getColor(value(0))
    ElseIf Me.day5 = "" And Me.Tag = "cust" Then
        Me.weekday5.Visible = False
        Me.task5 = ""
        Me.time5 = ""
    Else
        Me.day5 = ""
        Me.weekday5.Visible = False
    End If
    If Me.weekday5 = "ERR OR" Then
        Me.day5 = ""
        Me.weekday5.Visible = False
    End If
End Sub

Private Sub day6_Change()
    Dim timee As Integer, usedate As Date
    Dim value
    usedate = Sheet13.Range("E2").value
    'value holds the following:'
        'value(0) = weekcolor'
        'value(1) = weekday'
        'value(6) = todo'
    If IsNumeric(day6.Text) And IsEmpty(day6.Text) = False Then
        Me.weekday6.Visible = True
        value = GetDay(usedate, CInt(Me.day6))
        Me.weekday6 = value(0) & " " & value(1)
        Me.weekday6.BackColor = getColor(value(0))
    ElseIf Me.day6 = "" And Me.Tag = "cust" Then
        Me.weekday6.Visible = False
        Me.task6 = ""
        Me.time6 = ""
    Else
        Me.day6 = ""
        Me.weekday6.Visible = False
    End If
    If Me.weekday6 = "ERR OR" Then
        Me.day6 = ""
        Me.weekday6.Visible = False
    End If
End Sub

Private Sub day7_Change()
    Dim timee As Integer, usedate As Date
    Dim value
    usedate = Sheet13.Range("E2").value
    'value holds the following:'
        'value(0) = weekcolor'
        'value(1) = weekday'
        'value(7) = todo'
    If IsNumeric(day7.Text) And IsEmpty(day7.Text) = False Then
        Me.weekday7.Visible = True
        value = GetDay(usedate, CInt(Me.day7))
        Me.weekday7 = value(0) & " " & value(1)
        Me.weekday7.BackColor = getColor(value(0))
    ElseIf Me.day7 = "" And Me.Tag = "cust" Then
        Me.weekday7.Visible = False
        Me.task7 = ""
        Me.time7 = ""
    Else
        Me.day7 = ""
        Me.weekday7.Visible = False
    End If
    If Me.weekday7 = "ERR OR" Then
        Me.day7 = ""
        Me.weekday7.Visible = False
    End If
End Sub


'''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''
'''''''''Submission Button  ''''''''''
'''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''

Private Sub submit_Click()
    Dim globalyear As String
    globalyear = Sheet1.Range("K3").value
    Dim d As Integer
    Dim t As Integer
    Dim det, timee
    For d = 1 To 7
        If Controls.Item(Replace("weekday" & Str(d), " ", "")).Visible = False Or Controls.Item(Replace("task" & Str(d), " ", "")).Text = "" Then
            'do nothing, dont add task'
        Else
            det = Split(Controls.Item(Replace("weekday" & Str(d), " ", "")), " ")(1) & "/" & CStr(globalyear)
            If Controls.Item(Replace("time" & Str(d), " ", "")) = "*" Then
                'do nothing'
                timee = "*"
            Else
                timee = Replace(Controls.Item(Replace("time" & Str(d), " ", "")), " mins", "")
                If timee = "" Or IsNumeric(timee) = False Then
                    timee = 0
                End If
                timee = Replace(CStr(timee), " ", "")
            End If
            'AddTask'
            'Day 3 Read (time) Chr(128) 3/31/2021 Chr(138) notes_optional'
            If InStr(Controls.Item(Replace("task" & Str(d), " ", "")), "//") Then
                Dim task, note As String
                task = Trim(Split(Controls.Item(Replace("task" & Str(d), " ", "")), "//")(0))
                note = Trim(Split(Controls.Item(Replace("task" & Str(d), " ", "")), "//")(1))
                If timee = "*" Then
                    AddTask (task & Chr(128) & det & Chr(138) & note)
                Else
                    AddTask (task & " (" & timee & ")" & Chr(128) & det & Chr(138) & note)
                End If
                
                t = t + 1
            Else
                If timee = "*" Then
                    AddTask (Trim(Controls.Item(Replace("task" & Str(d), " ", ""))) & Chr(128) & det & Chr(138))
                Else
                    AddTask (Trim(Controls.Item(Replace("task" & Str(d), " ", ""))) & " (" & timee & ")" & Chr(128) & det & Chr(138))
                End If
                t = t + 1
            End If
        End If
    Next
    MsgBox ("Added " & CStr(t) & " tasks to BB..." & Chr(10) & "Do not forget to make your tapes!")
    Unload Me
End Sub

'''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''
'''Prevent "(" ")" character in task''''''
'''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''

Private Sub task1_Change()
    Dim oldtask
    oldtask = task1.Text
    If InStr(task1.Text, "(") Then
        task1 = Replace(oldtask, "(", "")
    End If
    If InStr(task1.Text, ")") Then
        task1 = Replace(oldtask, ")", "")
    End If
End Sub

Private Sub task2_Change()
    Dim oldtask
    oldtask = task2.Text
    If InStr(task2.Text, "(") Then
        task2 = Replace(oldtask, "(", "")
    End If
    If InStr(task2.Text, ")") Then
        task2 = Replace(oldtask, ")", "")
    End If
End Sub

Private Sub task3_Change()
    Dim oldtask
    oldtask = task3.Text
    If InStr(task3.Text, "(") Then
        task3 = Replace(oldtask, "(", "")
    End If
    If InStr(task3.Text, ")") Then
        task3 = Replace(oldtask, ")", "")
    End If
End Sub

Private Sub task4_Change()
    Dim oldtask
    oldtask = task4.Text
    If InStr(task4.Text, "(") Then
        task4 = Replace(oldtask, "(", "")
    End If
    If InStr(task4.Text, ")") Then
        task4 = Replace(oldtask, ")", "")
    End If
End Sub

Private Sub task5_Change()
    Dim oldtask
    oldtask = task5.Text
    If InStr(task5.Text, "(") Then
        task5 = Replace(oldtask, "(", "")
    End If
    If InStr(task5.Text, ")") Then
        task5 = Replace(oldtask, ")", "")
    End If
End Sub

Private Sub task6_Change()
    Dim oldtask
    oldtask = task6.Text
    If InStr(task6.Text, "(") Then
        task6 = Replace(oldtask, "(", "")
    End If
    If InStr(task6.Text, ")") Then
        task6 = Replace(oldtask, ")", "")
    End If
End Sub

Private Sub task7_Change()
    Dim oldtask
    oldtask = task7.Text
    If InStr(task7.Text, "(") Then
        task7 = Replace(oldtask, "(", "")
    End If
    If InStr(task7.Text, ")") Then
        task7 = Replace(oldtask, ")", "")
    End If
End Sub

'''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''
'''Find time for reading  ''''''
'''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''

Private Sub textNumberof_Change()
    Dim timee As Integer
    If IsNumeric(textNumberof.Text) And IsEmpty(textNumberof.Text) = False Then
        timee = CInt(textNumberof.Text)
        Select Case Me.Tag
        Case "mf"
            Me.time1.Text = CStr(2) & " mins"
            Me.time2.Text = CStr(2) & " mins"
            Me.time3.Text = CStr(getTime(timee, Me.Tag)) & " mins"
            Me.time4.Text = CStr(getTime(timee, Me.Tag)) & " mins"
            Me.time5.Text = CStr(getTime(timee, Me.Tag) + 20) & " mins"
        Case "dt"
            Me.time1.Text = CStr(2) & " mins"
            Me.time2.Text = CStr(2) & " mins"
            Me.time3.Text = CStr(getTime(timee, Me.Tag)) & " mins"
            Me.time4.Text = CStr(getTime(timee, Me.Tag)) & " mins"
            Me.time5.Text = CStr(getTime(timee, Me.Tag) + 20) & " mins"
        Case "noniso"
            Me.time1.Text = CStr(1) & " mins"
            Me.time2.Text = CStr(1) & " mins"
            Me.time3.Text = CStr(getTime(timee, Me.Tag)) & " mins"
            Me.time4.Text = CStr(getTime(timee, Me.Tag)) & " mins"
            Me.time5.Text = CStr(getTime(timee, Me.Tag) + 20) & " mins"
        Case "dtt"
            Me.time1.Text = CStr(2) & " mins"
            Me.time2.Text = CStr(2) & " mins"
            Me.time3.Text = CStr(getTime(timee, Me.Tag)) & " mins"
        Case "bfmf"
            Me.time1.Text = CStr(10) & " mins"
            Me.time2.Text = CStr(20) & " mins"
            Me.time3.Text = CStr(10) & " mins"
        Case "bfdt"
            Me.time1.Text = CStr(10) & " mins"
            Me.time2.Text = CStr(20) & " mins"
            Me.time3.Text = CStr(10) & " mins"
        End Select
    ElseIf textNumberof.Text = "" Then
        'do nothing'
    Else
        textNumberof.Text = ""
        MsgBox ("Thats not a number... Integers only please!")
    End If
End Sub

'''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''
'''Changing time to time " mins"'''
'''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''


Private Sub time1_AfterUpdate()
    If time1 = "*" Then
        'do nothing'
    ElseIf IsNumeric(time1) Then
        time1 = time1 & " mins"
    ElseIf IsNumeric(Replace(time1, "mins", "")) Then
        time1 = Trim(Replace(time1, "mins", "")) & " mins"
    Else
        time1 = ""
    End If
End Sub


Private Sub time2_AfterUpdate()
    If time2 = "*" Then
        'do nothing'
    ElseIf IsNumeric(time2) Then
        time2 = time2 & " mins"
    ElseIf IsNumeric(Replace(time2, "mins", "")) Then
        time2 = Trim(Replace(time2, "mins", "")) & " mins"
    Else
        time2 = ""
    End If
End Sub

Private Sub time3_AfterUpdate()
    If time3 = "*" Then
        'do nothing'
    ElseIf IsNumeric(time3) Then
        time3 = time3 & " mins"
    ElseIf IsNumeric(Replace(time3, "mins", "")) Then
        time3 = Trim(Replace(time3, "mins", "")) & " mins"
    Else
        time3 = ""
    End If
End Sub

Private Sub time4_AfterUpdate()
    If time4 = "*" Then
        'do nothing'
    ElseIf IsNumeric(time4) Then
        time4 = time4 & " mins"
    ElseIf IsNumeric(Replace(time4, "mins", "")) Then
        time4 = Trim(Replace(time4, "mins", "")) & " mins"
    Else
        time4 = ""
    End If
End Sub

Private Sub time5_AfterUpdate()
    If time5 = "*" Then
        'do nothing'
    ElseIf IsNumeric(time5) Then
        time5 = time5 & " mins"
    ElseIf IsNumeric(Replace(time5, "mins", "")) Then
        time5 = Trim(Replace(time5, "mins", "")) & " mins"
    Else
        time5 = ""
    End If
End Sub

Private Sub time6_AfterUpdate()
    If time6 = "*" Then
        'do nothing'
    ElseIf IsNumeric(time6) Then
        time6 = time6 & " mins"
    ElseIf IsNumeric(Replace(time6, "mins", "")) Then
        time6 = Trim(Replace(time6, "mins", "")) & " mins"
    Else
        time6 = ""
    End If
End Sub

Private Sub time7_AfterUpdate()
    If time7 = "*" Then
        'do nothing'
    ElseIf IsNumeric(time7) Then
        time7 = time7 & " mins"
    ElseIf IsNumeric(Replace(time7, "mins", "")) Then
        time7 = Trim(Replace(time7, "mins", "")) & " mins"
    Else
        time7 = ""
    End If
End Sub



Private Sub UserForm_Initialize()
    Call StopTimer
    Call SetTimer
End Sub

'''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''
'''''Changing Pending Day ''''''''''
'''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''

Private Sub weekday1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.labelPending.Caption = Me.weekday1
    Me.labelPending.BackColor = Me.weekday1.BackColor
    Me.labelPending.Tag = Me.day1
    Call updatePend
End Sub

Private Sub weekday2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.labelPending.Caption = Me.weekday2
    Me.labelPending.BackColor = Me.weekday2.BackColor
    Me.labelPending.Tag = Me.day2
    Call updatePend
End Sub

Private Sub weekday3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.labelPending.Caption = Me.weekday3
    Me.labelPending.BackColor = Me.weekday3.BackColor
    Me.labelPending.Tag = Me.day3
    Call updatePend
End Sub

Private Sub weekday4_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.labelPending.Caption = Me.weekday4
    Me.labelPending.BackColor = Me.weekday4.BackColor
    Me.labelPending.Tag = Me.day4
    Call updatePend
End Sub

Private Sub weekday5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.labelPending.Caption = Me.weekday5
    Me.labelPending.BackColor = Me.weekday5.BackColor
    Me.labelPending.Tag = Me.day5
    Call updatePend
End Sub

Private Sub weekday6_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.labelPending.Caption = Me.weekday6
    Me.labelPending.BackColor = Me.weekday6.BackColor
    Me.labelPending.Tag = Me.day6
    Call updatePend
End Sub

Private Sub weekday7_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.labelPending.Caption = Me.weekday7
    Me.labelPending.BackColor = Me.weekday7.BackColor
    Me.labelPending.Tag = Me.day7
    Call updatePend
End Sub

Private Sub UserForm_Initialize()
    Call StopTimer
    Call SetTimer
End Sub

Private Sub buttonSubmit_Click()
    'Day 3 Read Chr(128) 3/31/2021 Chr(138) notes_optional'
    ActiveSheet.Unprotect
    Dim note
    note = Trim(Me.textNote)
    note = Replace(note, ")", "")
    note = Replace(note, "(", "")
    note = Replace(note, ">", "")
    note = Replace(note, "<", "")
    If Me.textNote = "" Then
        Exit Sub
    ElseIf textTime <> "" Then
        AddTask (note & " (" & textTime.Text & ")" & Chr(128) & Me.textDate & Chr(138) & Trim(textNoteComment.Text))
    Else
        AddTask (note & Chr(128) & Me.textDate & Chr(138) & Trim(textNoteComment.Text))
    End If
    Unload Me
    ActiveSheet.Protect
End Sub

Private Sub textNote_Change()
    textNote.Text = Replace(textNote.Text, "(", "")
    textNote.Text = Replace(textNote.Text, ")", "")
    textNote.Text = Replace(textNote.Text, "<", "")
    textNote.Text = Replace(textNote.Text, ">", "")
End Sub

Private Sub textNoteComment_Change()
    textNoteComment.Text = Replace(textNoteComment.Text, "(", "")
    textNoteComment.Text = Replace(textNoteComment.Text, ")", "")
    textNoteComment.Text = Replace(textNoteComment.Text, "<", "")
    textNoteComment.Text = Replace(textNoteComment.Text, ">", "")
End Sub

Private Sub textTime_Change()
    textTime.Text = Trim(textTime.Text)
    If IsNumeric(textTime.Text) And InStr(textTime.Text, ".") = 0 Then
        If CInt(textTime.Text) < 1000 Then
            'do nothing'
        Else
            textTime.Text = ""
        End If
    Else
        textTime.Text = ""
    End If
End Sub

Private Sub UserForm_Initialize()
    Call StopTimer
    Call SetTimer
End Sub

Private Sub buttonSubmit_Click()
    'Day 3 Read Chr(128) 3/31/2021 Chr(138) notes_optional'
    ActiveSheet.Unprotect
    Dim note
    note = Trim(Me.textNote)
    note = Replace(note, ")", "")
    note = Replace(note, "(", "")
    note = Replace(note, ">", "")
    note = Replace(note, "<", "")
    If Me.textNote = "" Then
        Exit Sub
    ElseIf textTime <> "" Then
        AddTask (note & " (" & textTime.Text & ")" & Chr(128) & Me.textDate & Chr(138) & Trim(textNoteComment.Text))
    Else
        AddTask (note & Chr(128) & Me.textDate & Chr(138) & Trim(textNoteComment.Text))
    End If
    Unload Me
    ActiveSheet.Protect
End Sub

Private Sub textNote_Change()
    textNote.Text = Replace(textNote.Text, "(", "")
    textNote.Text = Replace(textNote.Text, ")", "")
    textNote.Text = Replace(textNote.Text, "<", "")
    textNote.Text = Replace(textNote.Text, ">", "")
End Sub

Private Sub textNoteComment_Change()
    textNoteComment.Text = Replace(textNoteComment.Text, "(", "")
    textNoteComment.Text = Replace(textNoteComment.Text, ")", "")
    textNoteComment.Text = Replace(textNoteComment.Text, "<", "")
    textNoteComment.Text = Replace(textNoteComment.Text, ">", "")
End Sub

Private Sub textTime_Change()
    textTime.Text = Trim(textTime.Text)
    If IsNumeric(textTime.Text) And InStr(textTime.Text, ".") = 0 Then
        If CInt(textTime.Text) < 1000 Then
            'do nothing'
        Else
            textTime.Text = ""
        End If
    Else
        textTime.Text = ""
    End If
End Sub










