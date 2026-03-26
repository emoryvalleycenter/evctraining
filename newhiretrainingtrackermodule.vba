Option Explicit

' ============================================================
' MONTHLY NEW HIRE TRACKER — Module2
'
' TRANSFER COLUMN MAP (header row 58, data rows 59-108):
'   A(1)=#  B(2)=Dept  C(3)=Last  D(4)=First
'   E(5)=From  F(6)=To  G(7)=MCF Received  H(8)=Effective Date
'   I(9)=Lift Van  J(10)=Job Desc/MOU  K(11)=UKERU  L(12)=Mealtime
'   M(13)=Delegations  N(14)=ITSP  O(15)=Therapies  P(16)=Status
'
' NEW HIRE COLUMN MAP (header row 4, data rows 5-54):
'   A(1)=#  B(2)=Dept  C(3)=Last  D(4)=First  E(5)=Bkgrd
'   F(6)=DOH  G(7)=Location/Title  H(8)=Assigned
'   I(9)=Relias  J(10)=3 Phase  K(11)=Job Desc  L(12)=CPR/FA
'   M(13)=Med Cert  N(14)=UKERU  O(15)=Mealtime  P(16)=Therapy
'   Q(17)=ITSP  R(18)=Delegation  S(19)=Status
' ============================================================

Sub GreyOutInactiveRows(ws As Worksheet)
    Dim r As Long, c As Long, statusVal As String
    For r = 5 To 54
        statusVal = Trim(CStr(ws.Cells(r, 19).Value))
        If statusVal <> "" And statusVal <> "Active" Then
            For c = 1 To 19
                ws.Cells(r, c).Interior.Color = RGB(217, 217, 217)
                ws.Cells(r, c).Font.Color = RGB(140, 140, 140)
                ws.Cells(r, c).Font.Strikethrough = True
            Next c
        Else
            For c = 1 To 19
                ws.Cells(r, c).Font.Color = RGB(51, 51, 51)
                ws.Cells(r, c).Font.Strikethrough = False
                If r Mod 2 = 0 Then
                    ws.Cells(r, c).Interior.Color = RGB(234, 240, 249)
                Else
                    ws.Cells(r, c).Interior.ColorIndex = xlNone
                End If
            Next c
        End If
    Next r
End Sub

Sub GreyOutTransferRows(ws As Worksheet)
    Dim r As Long, c As Long, statusVal As String
    For r = 59 To 108
        statusVal = UCase(Trim(CStr(ws.Cells(r, 16).Value)))
        If statusVal = "QUIT" Or statusVal = "TERMINATED" Or statusVal = "RESIGNED" Or statusVal = "NCNS" Then
            For c = 1 To 16
                ws.Cells(r, c).Interior.Color = RGB(217, 217, 217)
                ws.Cells(r, c).Font.Color = RGB(140, 140, 140)
                ws.Cells(r, c).Font.Strikethrough = True
            Next c
        Else
            For c = 1 To 16
                ws.Cells(r, c).Font.Color = RGB(51, 51, 51)
                ws.Cells(r, c).Font.Strikethrough = False
                If r Mod 2 = 0 Then
                    ws.Cells(r, c).Interior.Color = RGB(234, 240, 249)
                Else
                    ws.Cells(r, c).Interior.ColorIndex = xlNone
                End If
            Next c
        End If
    Next r
End Sub

Sub GreyOutAllSheets()
    Dim months As Variant, m As Long
    months = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
    Application.ScreenUpdating = False
    For m = 0 To 11
        GreyOutInactiveRows ThisWorkbook.Sheets(CStr(months(m)))
        GreyOutTransferRows ThisWorkbook.Sheets(CStr(months(m)))
    Next m
    Application.ScreenUpdating = True
    MsgBox "All inactive rows greyed out across all months.", vbInformation, "Done"
End Sub

' ============================================================
'  EnsureReportSheets - Creates Onboarding Report and/or
'  Termination Report sheets if they don't already exist,
'  with headers and formatting matching EVC tracker branding.
' ============================================================
Sub EnsureReportSheets()
    Dim wsOnb As Worksheet, wsTerm As Worksheet
    Dim c As Long
    Dim dash As String: dash = " " & ChrW(8212) & " "   ' em dash
    Dim arrow As String: arrow = ChrW(8592) & " "        ' left arrow
    
    Dim onbHeaders As Variant, termHeaders As Variant
    onbHeaders = Array("Month", "Dept", "Last Name", "First Name", "DOH", "Location / Title", "Missing Training", "Days Since Hire")
    termHeaders = Array("Month", "Dept", "Last Name", "First Name", "DOH", "Location / Title", "Status", "Notes")
    
    ' ===========================================
    '  ONBOARDING REPORT
    ' ===========================================
    On Error Resume Next: Set wsOnb = ThisWorkbook.Worksheets("Onboarding Report"): On Error GoTo 0
    If wsOnb Is Nothing Then
        Set wsOnb = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        wsOnb.Name = "Onboarding Report"
        
        ActiveWindow.DisplayGridlines = False
        
        wsOnb.Cells.Font.Name = "Aptos"
        wsOnb.Cells.Font.Size = 10
        wsOnb.Cells.Font.Color = RGB(51, 51, 51)
        
        wsOnb.Range("A1:H1").Merge
        With wsOnb.Range("A1")
            .Value = "ONBOARDING REPORT" & dash & "Incomplete Training"
            .Font.Name = "Grandview": .Font.Size = 16: .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(27, 42, 74)
            .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
            .IndentLevel = 1
        End With
        wsOnb.Rows(1).RowHeight = 33.75
        
        wsOnb.Range("A2").Value = arrow & "Dashboard"
        With wsOnb.Range("A2")
            .Font.Name = "Aptos": .Font.Size = 9
            .Font.Color = RGB(46, 80, 144)
            .Font.Underline = xlUnderlineStyleSingle
        End With
        wsOnb.Rows(2).RowHeight = 18
        
        For c = 1 To 8
            wsOnb.Cells(3, c).Value = onbHeaders(c - 1)
            With wsOnb.Cells(3, c)
                .Font.Name = "Grandview": .Font.Size = 9: .Font.Bold = True
                .Font.Color = RGB(255, 255, 255)
                .Interior.Color = RGB(46, 80, 144)
                .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Color = RGB(27, 42, 74)
                .Borders(xlEdgeBottom).Weight = xlMedium
            End With
        Next c
        wsOnb.Rows(3).RowHeight = 27.75
        
        wsOnb.Columns("A").ColumnWidth = 12
        wsOnb.Columns("B").ColumnWidth = 15
        wsOnb.Columns("C").ColumnWidth = 14
        wsOnb.Columns("D").ColumnWidth = 14
        wsOnb.Columns("E").ColumnWidth = 12
        wsOnb.Columns("F").ColumnWidth = 24
        wsOnb.Columns("G").ColumnWidth = 42
        wsOnb.Columns("H").ColumnWidth = 16
        
        wsOnb.Range("A4").Select
        ActiveWindow.FreezePanes = True
        
        With wsOnb.PageSetup
            .Orientation = xlLandscape
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .PrintTitleRows = "$1:$3"
        End With
    End If
    
    ' ===========================================
    '  TERMINATION REPORT
    ' ===========================================
    On Error Resume Next: Set wsTerm = ThisWorkbook.Worksheets("Termination Report"): On Error GoTo 0
    If wsTerm Is Nothing Then
        Set wsTerm = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        wsTerm.Name = "Termination Report"
        
        ActiveWindow.DisplayGridlines = False
        
        wsTerm.Cells.Font.Name = "Aptos"
        wsTerm.Cells.Font.Size = 10
        wsTerm.Cells.Font.Color = RGB(51, 51, 51)
        
        wsTerm.Range("A1:H1").Merge
        With wsTerm.Range("A1")
            .Value = "TERMINATION REPORT" & dash & "Separations"
            .Font.Name = "Grandview": .Font.Size = 16: .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(27, 42, 74)
            .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
            .IndentLevel = 1
        End With
        wsTerm.Rows(1).RowHeight = 33.75
        
        wsTerm.Range("A2").Value = arrow & "Dashboard"
        With wsTerm.Range("A2")
            .Font.Name = "Aptos": .Font.Size = 9
            .Font.Color = RGB(46, 80, 144)
            .Font.Underline = xlUnderlineStyleSingle
        End With
        wsTerm.Rows(2).RowHeight = 18
        
        For c = 1 To 8
            wsTerm.Cells(3, c).Value = termHeaders(c - 1)
            With wsTerm.Cells(3, c)
                .Font.Name = "Grandview": .Font.Size = 9: .Font.Bold = True
                .Font.Color = RGB(255, 255, 255)
                .Interior.Color = RGB(139, 46, 49)
                .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Color = RGB(90, 30, 30)
                .Borders(xlEdgeBottom).Weight = xlMedium
            End With
        Next c
        wsTerm.Rows(3).RowHeight = 27.75
        
        wsTerm.Columns("A").ColumnWidth = 12
        wsTerm.Columns("B").ColumnWidth = 15
        wsTerm.Columns("C").ColumnWidth = 14
        wsTerm.Columns("D").ColumnWidth = 14
        wsTerm.Columns("E").ColumnWidth = 12
        wsTerm.Columns("F").ColumnWidth = 24
        wsTerm.Columns("G").ColumnWidth = 14
        wsTerm.Columns("H").ColumnWidth = 24
        
        wsTerm.Range("A4").Select
        ActiveWindow.FreezePanes = True
        
        With wsTerm.PageSetup
            .Orientation = xlLandscape
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .PrintTitleRows = "$1:$3"
        End With
    End If
End Sub

' ============================================================
'  RefreshReports - Ensures sheets exist, then clears data
'  area and repopulates from all 12 monthly sheets.
' ============================================================
Sub RefreshReports()
    Application.ScreenUpdating = False
    Dim months As Variant, m As Long, r As Long, c As Long
    months = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
    For m = 0 To 11
        GreyOutInactiveRows ThisWorkbook.Sheets(CStr(months(m)))
        GreyOutTransferRows ThisWorkbook.Sheets(CStr(months(m)))
    Next m
    
    EnsureReportSheets
    
    ' --- ONBOARDING REPORT ---
    Dim wsOnb As Worksheet: Set wsOnb = ThisWorkbook.Sheets("Onboarding Report")
    wsOnb.Range("A4:H500").ClearContents
    wsOnb.Range("A4:H500").Interior.ColorIndex = xlNone
    Dim onbRow As Long: onbRow = 4
    Dim wsMonth As Worksheet
    Dim missing As String
    For m = 0 To 11
        Set wsMonth = ThisWorkbook.Sheets(CStr(months(m)))
        For r = 5 To 54
            If wsMonth.Cells(r, 3).Value <> "" And wsMonth.Cells(r, 19).Value = "Active" Then
                missing = ""
                If wsMonth.Cells(r, 9).Value <> "Yes" And wsMonth.Cells(r, 9).Value <> "N/A" Then missing = missing & "Relias; "
                If wsMonth.Cells(r, 10).Value <> "Yes" And wsMonth.Cells(r, 10).Value <> "N/A" Then missing = missing & "3 Phase; "
                If wsMonth.Cells(r, 11).Value <> "Yes" And wsMonth.Cells(r, 11).Value <> "N/A" Then missing = missing & "Job Desc; "
                If wsMonth.Cells(r, 12).Value <> "Yes" And wsMonth.Cells(r, 12).Value <> "N/A" Then missing = missing & "CPR/FA; "
                If wsMonth.Cells(r, 13).Value <> "Yes" And wsMonth.Cells(r, 13).Value <> "N/A" Then missing = missing & "Med Cert; "
                If wsMonth.Cells(r, 14).Value <> "Yes" And wsMonth.Cells(r, 14).Value <> "N/A" Then missing = missing & "UKERU; "
                If wsMonth.Cells(r, 15).Value <> "Yes" And wsMonth.Cells(r, 15).Value <> "N/A" Then missing = missing & "Mealtime; "
                If wsMonth.Cells(r, 16).Value <> "Yes" And wsMonth.Cells(r, 16).Value <> "N/A" Then missing = missing & "Therapy; "
                If wsMonth.Cells(r, 17).Value <> "Yes" And wsMonth.Cells(r, 17).Value <> "N/A" Then missing = missing & "ITSP; "
                If wsMonth.Cells(r, 18).Value <> "Yes" And wsMonth.Cells(r, 18).Value <> "N/A" Then missing = missing & "Delegation; "
                If missing <> "" Then
                    missing = Left(missing, Len(missing) - 2)
                    wsOnb.Cells(onbRow, 1).Value = CStr(months(m))
                    wsOnb.Cells(onbRow, 2).Value = wsMonth.Cells(r, 2).Value
                    wsOnb.Cells(onbRow, 3).Value = wsMonth.Cells(r, 3).Value
                    wsOnb.Cells(onbRow, 4).Value = wsMonth.Cells(r, 4).Value
                    wsOnb.Cells(onbRow, 5).Value = wsMonth.Cells(r, 6).Value
                    wsOnb.Cells(onbRow, 5).NumberFormat = "MM/DD/YYYY"
                    wsOnb.Cells(onbRow, 6).Value = wsMonth.Cells(r, 7).Value
                    wsOnb.Cells(onbRow, 7).Value = missing
                    If IsDate(wsMonth.Cells(r, 6).Value) Then
                        wsOnb.Cells(onbRow, 8).Value = Date - CDate(wsMonth.Cells(r, 6).Value)
                    End If
                    For c = 1 To 8
                        wsOnb.Cells(onbRow, c).Font.Name = "Aptos"
                        wsOnb.Cells(onbRow, c).Font.Size = 10
                        wsOnb.Cells(onbRow, c).Borders.LineStyle = xlContinuous
                        wsOnb.Cells(onbRow, c).Borders.Color = RGB(208, 208, 208)
                    Next c
                    If onbRow Mod 2 = 0 Then
                        wsOnb.Range(wsOnb.Cells(onbRow, 1), wsOnb.Cells(onbRow, 8)).Interior.Color = RGB(234, 240, 249)
                    End If
                    onbRow = onbRow + 1
                End If
            End If
        Next r
    Next m
    
    ' --- TERMINATION REPORT ---
    Dim wsTerm As Worksheet: Set wsTerm = ThisWorkbook.Sheets("Termination Report")
    wsTerm.Range("A4:H500").ClearContents
    wsTerm.Range("A4:H500").Interior.ColorIndex = xlNone
    Dim termRow As Long: termRow = 4
    Dim stat As String
    For m = 0 To 11
        Set wsMonth = ThisWorkbook.Sheets(CStr(months(m)))
        For r = 5 To 54
            If wsMonth.Cells(r, 3).Value <> "" Then
                stat = CStr(wsMonth.Cells(r, 19).Value)
                If stat = "Terminated" Or stat = "Resigned" Or stat = "NCNS" Or stat = "Suspended" Then
                    wsTerm.Cells(termRow, 1).Value = CStr(months(m))
                    wsTerm.Cells(termRow, 2).Value = wsMonth.Cells(r, 2).Value
                    wsTerm.Cells(termRow, 3).Value = wsMonth.Cells(r, 3).Value
                    wsTerm.Cells(termRow, 4).Value = wsMonth.Cells(r, 4).Value
                    wsTerm.Cells(termRow, 5).Value = wsMonth.Cells(r, 6).Value
                    wsTerm.Cells(termRow, 5).NumberFormat = "MM/DD/YYYY"
                    wsTerm.Cells(termRow, 6).Value = wsMonth.Cells(r, 7).Value
                    wsTerm.Cells(termRow, 7).Value = stat
                    wsTerm.Cells(termRow, 8).Value = ""
                    For c = 1 To 8
                        wsTerm.Cells(termRow, c).Font.Name = "Aptos"
                        wsTerm.Cells(termRow, c).Font.Size = 10
                        wsTerm.Cells(termRow, c).Borders.LineStyle = xlContinuous
                        wsTerm.Cells(termRow, c).Borders.Color = RGB(208, 208, 208)
                    Next c
                    If termRow Mod 2 = 0 Then
                        wsTerm.Range(wsTerm.Cells(termRow, 1), wsTerm.Cells(termRow, 8)).Interior.Color = RGB(253, 232, 232)
                    End If
                    termRow = termRow + 1
                End If
            End If
        Next r
    Next m
    Application.ScreenUpdating = True
    MsgBox "Reports refreshed!" & vbCrLf & _
           "Incomplete training: " & (onbRow - 4) & " employees" & vbCrLf & _
           "Terminations: " & (termRow - 4) & " employees", vbInformation, "Reports Updated"
End Sub

Sub ShowIncompleteTraining()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim months As Variant, isMonth As Boolean, m As Long, r As Long, c As Long, count As Long
    months = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
    isMonth = False
    For m = 0 To 11
        If ws.Name = CStr(months(m)) Then isMonth = True: Exit For
    Next m
    If Not isMonth Then MsgBox "Navigate to a month sheet first.", vbInformation: Exit Sub
    GreyOutInactiveRows ws
    count = 0
    For r = 5 To 54
        If ws.Cells(r, 3).Value <> "" And ws.Cells(r, 19).Value = "Active" Then
            For c = 9 To 18
                If ws.Cells(r, c).Value <> "Yes" And ws.Cells(r, c).Value <> "N/A" Then
                    ws.Cells(r, c).Interior.Color = RGB(255, 200, 200)
                    count = count + 1
                End If
            Next c
        End If
    Next r
    MsgBox count & " incomplete items highlighted in red.", vbInformation, "Training Check"
End Sub

Sub GoToOnboardingReport()
    EnsureReportSheets
    Call RefreshReports
    ThisWorkbook.Sheets("Onboarding Report").Activate
End Sub

Sub GoToTerminationReport()
    EnsureReportSheets
    Call RefreshReports
    ThisWorkbook.Sheets("Termination Report").Activate
End Sub

Sub PrintDashboard()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Dashboard")
    ws.PageSetup.PrintArea = "A1:R52"
    ws.PageSetup.Orientation = xlLandscape
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = False
    ws.PrintPreview
End Sub

Sub AddAndProcessTraining()
    Dim wsResults As Worksheet, wsMonth As Worksheet
    Dim trainingSession As String, passOrFail As String, participantName As String
    Dim trainingDateVal As Variant, trainingDate As String, trainingCol As Long
    Dim fName As String, lName As String, parts() As String, missing As String
    Dim logRow As Long, foundSheet As String, foundRow As Long, foundIt As Boolean
    Dim skipReason As String, confirm As String, empDisplay As String
    Dim months As Variant, m As Long, mRow As Long
    Dim cellLast As String, cellFirst As String, fMatch As Boolean, empStatus As String
    Dim targetCell As Range
    On Error Resume Next: Set wsResults = ThisWorkbook.Worksheets("Training Results"): On Error GoTo 0
    If wsResults Is Nothing Then MsgBox "Cannot find 'Training Results' sheet.", vbExclamation: Exit Sub
    trainingSession = Trim(wsResults.Range("A5").Value & "")
    passOrFail = Trim(wsResults.Range("B5").Value & "")
    participantName = Trim(wsResults.Range("C5").Value & "")
    trainingDateVal = wsResults.Range("D5").Value
    missing = ""
    If trainingSession = "" Then missing = missing & "  - Training Session" & vbCrLf
    If passOrFail = "" Then missing = missing & "  - Result" & vbCrLf
    If participantName = "" Then missing = missing & "  - Participant Name" & vbCrLf
    If Trim(trainingDateVal & "") = "" Then missing = missing & "  - Date of Training" & vbCrLf
    If missing <> "" Then MsgBox "Please fill in all fields:" & vbCrLf & vbCrLf & missing, vbExclamation, "Missing Info": Exit Sub
    If IsDate(trainingDateVal) Then trainingDate = Format(trainingDateVal, "MM/DD/YYYY") Else trainingDate = Trim(CStr(trainingDateVal))
    trainingCol = TrainCol(trainingSession)
    If trainingCol = 0 Then MsgBox "Unknown training: '" & trainingSession & "'", vbExclamation: Exit Sub
    If InStr(participantName, ",") > 0 Then
        parts = Split(participantName, ","): lName = Trim(parts(0)): fName = Trim(parts(1))
    ElseIf InStr(participantName, " ") > 0 Then
        fName = Trim(Left(participantName, InStr(participantName, " ") - 1))
        lName = Trim(Mid(participantName, InStr(participantName, " ") + 1))
    Else: lName = participantName: fName = ""
    End If
    months = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
    foundIt = False: skipReason = ""
    For m = LBound(months) To UBound(months)
        Set wsMonth = Nothing: On Error Resume Next: Set wsMonth = ThisWorkbook.Worksheets(CStr(months(m))): On Error GoTo 0
        If wsMonth Is Nothing Then GoTo NextMQ
        For mRow = 5 To 54
            cellLast = Trim(wsMonth.Cells(mRow, 3).Value & "")
            cellFirst = Trim(wsMonth.Cells(mRow, 4).Value & "")
            If cellLast = "" Then GoTo NextRQ
            If StrComp(cellLast, lName, vbTextCompare) = 0 Then
                fMatch = False
                If fName = "" Then fMatch = True
                If StrComp(cellFirst, fName, vbTextCompare) = 0 Then fMatch = True
                If InStr(1, cellFirst, fName, vbTextCompare) > 0 Then fMatch = True
                If InStr(1, fName, cellFirst, vbTextCompare) > 0 Then fMatch = True
                If fMatch Then
                    empStatus = UCase(Trim(wsMonth.Cells(mRow, 19).Value & ""))
                    If empStatus = "TERMINATED" Or empStatus = "NCNS" Or empStatus = "RESIGNED" Then
                        skipReason = "Found on " & CStr(months(m)) & " but status is " & empStatus: GoTo NextRQ
                    End If
                    foundIt = True: foundSheet = CStr(months(m)): foundRow = mRow: GoTo DoneSQ
                End If
            End If
NextRQ:
        Next mRow
NextMQ:
    Next m
DoneSQ:
    Application.ScreenUpdating = False
    logRow = 9
    Do While logRow <= 508
        If Trim(wsResults.Cells(logRow, 3).Value & "") = "" And Trim(wsResults.Cells(logRow, 4).Value & "") = "" Then Exit Do
        logRow = logRow + 1
    Loop
    If logRow > 508 Then MsgBox "Log is full.", vbExclamation: Application.ScreenUpdating = True: Exit Sub
    wsResults.Cells(logRow, 1).Value = Date
    wsResults.Cells(logRow, 2).Value = passOrFail
    wsResults.Cells(logRow, 3).Value = trainingSession
    wsResults.Cells(logRow, 4).Value = participantName
    wsResults.Cells(logRow, 5).Value = trainingDateVal
    If foundIt Then
        Set wsMonth = ThisWorkbook.Worksheets(foundSheet)
        Set targetCell = wsMonth.Cells(foundRow, trainingCol)
        Select Case UCase(passOrFail)
            Case "PASS"
                targetCell.Value = "Yes"
                If Not targetCell.Comment Is Nothing Then targetCell.Comment.Delete
                targetCell.AddComment "Completed: " & trainingDate
                wsResults.Range(wsResults.Cells(logRow, 1), wsResults.Cells(logRow, 5)).Interior.Color = RGB(198, 239, 206)
            Case "FAIL"
                targetCell.Value = "No"
                If Not targetCell.Comment Is Nothing Then targetCell.Comment.Delete
                targetCell.AddComment "Failed: " & trainingDate
                wsResults.Range(wsResults.Cells(logRow, 1), wsResults.Cells(logRow, 5)).Interior.Color = RGB(255, 199, 206)
            Case "N/A"
                targetCell.Value = "N/A"
                wsResults.Range(wsResults.Cells(logRow, 1), wsResults.Cells(logRow, 5)).Interior.Color = RGB(255, 235, 156)
        End Select
        wsResults.Cells(logRow, 6).Value = "Yes"
    Else
        wsResults.Cells(logRow, 6).Value = "NOT FOUND"
        wsResults.Cells(logRow, 6).Font.Color = RGB(255, 0, 0)
        wsResults.Cells(logRow, 6).Font.Bold = True
    End If
    RefreshOnboardingReport
    wsResults.Range("A5:D5").ClearContents
    Application.ScreenUpdating = True
    If foundIt Then
        empDisplay = wsMonth.Cells(foundRow, 4).Value & " " & wsMonth.Cells(foundRow, 3).Value
        confirm = "Done!" & vbCrLf & vbCrLf & "  " & empDisplay & vbCrLf & _
                  "  " & trainingSession & " = " & UCase(passOrFail) & vbCrLf & _
                  "  Updated on " & foundSheet & " sheet"
        If UCase(passOrFail) = "PASS" Then confirm = confirm & vbCrLf & "  Completed " & trainingDate
    Else
        confirm = "Logged but employee not found on monthly sheets."
        If skipReason <> "" Then confirm = confirm & vbCrLf & skipReason
    End If
    MsgBox confirm, vbInformation, "Training Logged"
End Sub

Sub RefreshOnboardingReport()
    EnsureReportSheets
    
    Dim wsReport As Worksheet, wsMonth As Worksheet, months As Variant
    Dim m As Long, mRow As Long, reportRow As Long, c As Long
    Dim cellVal As String, missingItems As String, empLast As String, empStat As String
    Dim tCols As Variant, tNames As Variant, clearEnd As Long
    On Error Resume Next: Set wsReport = ThisWorkbook.Worksheets("Onboarding Report"): On Error GoTo 0
    If wsReport Is Nothing Then Exit Sub
    clearEnd = wsReport.Cells(wsReport.Rows.count, "A").End(xlUp).Row
    If clearEnd < 4 Then clearEnd = 4
    wsReport.Range("A4:H" & clearEnd + 5).ClearContents
    wsReport.Range("A4:H" & clearEnd + 5).Interior.ColorIndex = xlNone
    months = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
    tCols = Array(9, 10, 11, 12, 13, 14, 15, 16, 17, 18)
    tNames = Array("Relias", "3 Phase", "Job Desc", "CPR/FA", "Med Cert", "UKERU", "Mealtime", "Therapy", "ITSP", "Delegation")
    reportRow = 4
    For m = LBound(months) To UBound(months)
        Set wsMonth = Nothing: On Error Resume Next: Set wsMonth = ThisWorkbook.Worksheets(CStr(months(m))): On Error GoTo 0
        If wsMonth Is Nothing Then GoTo NextMR
        For mRow = 5 To 54
            empLast = Trim(wsMonth.Cells(mRow, 3).Value & "")
            If empLast = "" Then GoTo NextRR
            empStat = UCase(Trim(wsMonth.Cells(mRow, 19).Value & ""))
            If empStat = "TERMINATED" Or empStat = "NCNS" Or empStat = "RESIGNED" Then GoTo NextRR
            missingItems = ""
            For c = LBound(tCols) To UBound(tCols)
                cellVal = UCase(Trim(wsMonth.Cells(mRow, tCols(c)).Value & ""))
                If cellVal <> "YES" And cellVal <> "N/A" Then
                    If missingItems <> "" Then missingItems = missingItems & "; "
                    missingItems = missingItems & tNames(c)
                End If
            Next c
            If missingItems <> "" Then
                wsReport.Cells(reportRow, 1).Value = CStr(months(m))
                wsReport.Cells(reportRow, 2).Value = wsMonth.Cells(mRow, 2).Value
                wsReport.Cells(reportRow, 3).Value = wsMonth.Cells(mRow, 3).Value
                wsReport.Cells(reportRow, 4).Value = wsMonth.Cells(mRow, 4).Value
                wsReport.Cells(reportRow, 5).Value = wsMonth.Cells(mRow, 6).Value
                wsReport.Cells(reportRow, 5).NumberFormat = "MM/DD/YYYY"
                wsReport.Cells(reportRow, 6).Value = wsMonth.Cells(mRow, 7).Value
                wsReport.Cells(reportRow, 7).Value = missingItems
                If IsDate(wsMonth.Cells(mRow, 6).Value) Then
                    wsReport.Cells(reportRow, 8).Value = DateDiff("d", wsMonth.Cells(mRow, 6).Value, Date)
                End If
                For c = 1 To 8
                    wsReport.Cells(reportRow, c).Font.Name = "Aptos"
                    wsReport.Cells(reportRow, c).Font.Size = 10
                    wsReport.Cells(reportRow, c).Borders.LineStyle = xlContinuous
                    wsReport.Cells(reportRow, c).Borders.Color = RGB(208, 208, 208)
                Next c
                If reportRow Mod 2 = 0 Then
                    wsReport.Range(wsReport.Cells(reportRow, 1), wsReport.Cells(reportRow, 8)).Interior.Color = RGB(234, 240, 249)
                End If
                reportRow = reportRow + 1
            End If
NextRR:
        Next mRow
NextMR:
    Next m
End Sub

Sub FillInProgress()
    Dim wsMonth As Worksheet, months As Variant, m As Long, mRow As Long, col As Long, filled As Long
    Dim empStat As String
    Application.ScreenUpdating = False
    months = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
    For m = LBound(months) To UBound(months)
        Set wsMonth = Nothing: On Error Resume Next: Set wsMonth = ThisWorkbook.Worksheets(CStr(months(m))): On Error GoTo 0
        If wsMonth Is Nothing Then GoTo NextMF
        For mRow = 5 To 54
            If Trim(wsMonth.Cells(mRow, 3).Value & "") = "" Then GoTo NextRF
            empStat = UCase(Trim(wsMonth.Cells(mRow, 19).Value & ""))
            If empStat = "TERMINATED" Or empStat = "NCNS" Or empStat = "RESIGNED" Then GoTo NextRF
            For col = 9 To 18
                If Trim(wsMonth.Cells(mRow, col).Value & "") = "" Then
                    wsMonth.Cells(mRow, col).Value = "In Progress": filled = filled + 1
                End If
            Next col
NextRF:
        Next mRow
NextMF:
    Next m
    Application.ScreenUpdating = True
    MsgBox "Filled " & filled & " blank cells with 'In Progress'.", vbInformation
End Sub

Sub ResetProcessedFlags()
    Dim ws As Worksheet, lastRow As Long
    Set ws = ThisWorkbook.Worksheets("Training Results")
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    If lastRow < 9 Then Exit Sub
    If MsgBox("Clear all Processed flags and highlights?", vbYesNo + vbQuestion, "Reset") = vbYes Then
        ws.Range("F9:F" & lastRow).ClearContents
        ws.Range("A9:E" & lastRow).Interior.ColorIndex = xlNone
        MsgBox "Flags cleared.", vbInformation
    End If
End Sub

' ============================================================
' TrainCol — New Hire training column lookup (unchanged)
' ============================================================
Function TrainCol(ByVal s As String) As Long
    Select Case Trim(s)
        Case "Relias": TrainCol = 9
        Case "3 Phase": TrainCol = 10
        Case "Job Desc": TrainCol = 11
        Case "CPR/FA": TrainCol = 12
        Case "Med Cert": TrainCol = 13
        Case "UKERU": TrainCol = 14
        Case "Mealtime": TrainCol = 15
        Case "Therapy": TrainCol = 16
        Case "ITSP": TrainCol = 17
        Case "Delegation": TrainCol = 18
        Case Else: TrainCol = 0
    End Select
End Function

' ============================================================
' TrainColTransfer — Transfer training column lookup
' UPDATED for Effective Date insertion at col H(8)
'   I(9)=Lift Van  J(10)=Job Desc/MOU  K(11)=UKERU
'   L(12)=Mealtime  M(13)=Delegations  N(14)=ITSP
'   O(15)=Therapies
' ============================================================
Function TrainColTransfer(training As String) As Long
    Select Case UCase(Trim(training))
        Case "LIFT VAN", "VAN LIFT", "LIFT": TrainColTransfer = 9               ' I
        Case "JOB DESC", "JOB DESCRIPTION", "JOB DESC/MOU", "MOU": TrainColTransfer = 10  ' J
        Case "UKERU": TrainColTransfer = 11                                      ' K
        Case "MEALTIME": TrainColTransfer = 12                                   ' L
        Case "DELEGATIONS", "DELEGATION": TrainColTransfer = 13                  ' M
        Case "ITSP": TrainColTransfer = 14                                       ' N
        Case "THERAPIES", "THERAPY": TrainColTransfer = 15                       ' O
        Case Else: TrainColTransfer = 0
    End Select
End Function

Sub SyncTrainingResults()
    Dim wsResults As Worksheet, wsMonth As Worksheet
    Dim lastRow As Long, r As Long, m As Long, mRow As Long
    Dim passOrFail As String, trainingSession As String, participantName As String
    Dim trainingDate As String, trainingCol As Long, transferCol As Long
    Dim fName As String, lName As String, parts() As String
    Dim cellLast As String, cellFirst As String, fMatch As Boolean, empStatus As String
    Dim targetCell As Range, matchFound As Boolean, skipReason As String
    Dim totalProcessed As Long, totalPassed As Long, totalFailed As Long
    Dim totalNA As Long, totalNotFound As Long, totalSkipped As Long
    Dim logMessages As String, months As Variant
    
    On Error Resume Next: Set wsResults = ThisWorkbook.Worksheets("Training Results"): On Error GoTo 0
    If wsResults Is Nothing Then MsgBox "Cannot find 'Training Results' sheet.", vbExclamation: Exit Sub
    
    lastRow = wsResults.Cells(wsResults.Rows.count, "C").End(xlUp).Row
    If lastRow < 9 Then
        lastRow = wsResults.Cells(wsResults.Rows.count, "D").End(xlUp).Row
    End If
    If lastRow < 9 Then MsgBox "No results to sync.", vbInformation: Exit Sub
    
    Dim answer As VbMsgBoxResult
    answer = MsgBox("Sync " & (lastRow - 8) & " log rows to month sheets?" & vbCrLf & vbCrLf & _
                    "This will:" & vbCrLf & _
                    "  - Update training cells (Yes/No)" & vbCrLf & _
                    "  - Search BOTH New Hires and Transfers" & vbCrLf & _
                    "  - Add completion date as a note" & vbCrLf & _
                    "  - Skip already-processed rows" & vbCrLf & _
                    "  - Refresh the Onboarding Report", _
                    vbYesNo + vbQuestion, "Sync Training Results")
    If answer <> vbYes Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    months = Array("January", "February", "March", "April", "May", "June", _
                   "July", "August", "September", "October", "November", "December")
    
    For r = 9 To lastRow
        If UCase(Trim(wsResults.Cells(r, 6).Value & "")) = "YES" Then totalSkipped = totalSkipped + 1: GoTo NextRow
        
        trainingSession = Trim(wsResults.Cells(r, 3).Value & "")
        participantName = Trim(wsResults.Cells(r, 4).Value & "")
        If trainingSession = "" Or participantName = "" Then GoTo NextRow
        
        passOrFail = Trim(wsResults.Cells(r, 2).Value & "")
        If passOrFail = "" Then passOrFail = "Pass"
        
        If IsDate(wsResults.Cells(r, 5).Value) Then
            trainingDate = Format(wsResults.Cells(r, 5).Value, "MM/DD/YYYY")
        ElseIf Trim(wsResults.Cells(r, 5).Value & "") <> "" Then
            trainingDate = Trim(wsResults.Cells(r, 5).Value & "")
        ElseIf IsDate(wsResults.Cells(r, 1).Value) Then
            trainingDate = Format(wsResults.Cells(r, 1).Value, "MM/DD/YYYY")
        Else
            trainingDate = Format(Date, "MM/DD/YYYY")
        End If
        
        trainingCol = TrainCol(trainingSession)
        transferCol = TrainColTransfer(trainingSession)
        
        If trainingCol = 0 And transferCol = 0 Then
            logMessages = logMessages & "Row " & r & ": Unknown training '" & trainingSession & "'" & vbCrLf
            GoTo NextRow
        End If
        
        If InStr(participantName, ",") > 0 Then
            parts = Split(participantName, ","): lName = Trim(parts(0)): fName = Trim(parts(1))
        ElseIf InStr(participantName, " ") > 0 Then
            fName = Trim(Left(participantName, InStr(participantName, " ") - 1))
            lName = Trim(Mid(participantName, InStr(participantName, " ") + 1))
        Else: lName = participantName: fName = ""
        End If
        
        matchFound = False: skipReason = ""
        
        For m = LBound(months) To UBound(months)
            Set wsMonth = Nothing
            On Error Resume Next: Set wsMonth = ThisWorkbook.Worksheets(CStr(months(m))): On Error GoTo 0
            If wsMonth Is Nothing Then GoTo NextMonth
            
            ' -- NEW HIRES (Rows 5-54) --
            If trainingCol > 0 Then
                For mRow = 5 To 54
                    cellLast = Trim(wsMonth.Cells(mRow, 3).Value & "")
                    cellFirst = Trim(wsMonth.Cells(mRow, 4).Value & "")
                    If cellLast = "" Then GoTo NextMRow1
                    
                    If StrComp(cellLast, lName, vbTextCompare) = 0 Then
                        fMatch = False
                        If fName = "" Then fMatch = True
                        If StrComp(cellFirst, fName, vbTextCompare) = 0 Then fMatch = True
                        If InStr(1, cellFirst, fName, vbTextCompare) > 0 Then fMatch = True
                        If InStr(1, fName, cellFirst, vbTextCompare) > 0 Then fMatch = True
                        
                        If fMatch Then
                            empStatus = UCase(Trim(wsMonth.Cells(mRow, 19).Value & ""))
                            If empStatus = "TERMINATED" Or empStatus = "NCNS" Or empStatus = "RESIGNED" Then
                                skipReason = empStatus: GoTo NextMRow1
                            End If
                            
                            Set targetCell = wsMonth.Cells(mRow, trainingCol)
                            GoSub ApplyResult
                            matchFound = True
                            wsResults.Cells(r, 6).Value = "Yes"
                            totalProcessed = totalProcessed + 1
                            GoTo DoneSearch
                        End If
                    End If
NextMRow1:
                Next mRow
            End If
            
            ' -- TRANSFERS (Rows 59-108, Status at col 16) --
            If transferCol > 0 Then
                For mRow = 59 To 108
                    cellLast = Trim(wsMonth.Cells(mRow, 3).Value & "")
                    cellFirst = Trim(wsMonth.Cells(mRow, 4).Value & "")
                    If cellLast = "" Then GoTo NextMRow2
                    
                    If StrComp(cellLast, lName, vbTextCompare) = 0 Then
                        fMatch = False
                        If fName = "" Then fMatch = True
                        If StrComp(cellFirst, fName, vbTextCompare) = 0 Then fMatch = True
                        If InStr(1, cellFirst, fName, vbTextCompare) > 0 Then fMatch = True
                        If InStr(1, fName, cellFirst, vbTextCompare) > 0 Then fMatch = True
                        
                        If fMatch Then
                            empStatus = UCase(Trim(wsMonth.Cells(mRow, 16).Value & ""))
                            If empStatus = "TERMINATED" Or empStatus = "NCNS" Or empStatus = "RESIGNED" Or empStatus = "QUIT" Then
                                skipReason = empStatus: GoTo NextMRow2
                            End If
                            
                            Set targetCell = wsMonth.Cells(mRow, transferCol)
                            GoSub ApplyResult
                            matchFound = True
                            wsResults.Cells(r, 6).Value = "Yes"
                            totalProcessed = totalProcessed + 1
                            GoTo DoneSearch
                        End If
                    End If
NextMRow2:
                Next mRow
            End If
            
NextMonth:
        Next m
        
DoneSearch:
        If Not matchFound Then
            totalNotFound = totalNotFound + 1
            wsResults.Cells(r, 6).Value = "NOT FOUND"
            wsResults.Cells(r, 6).Font.Color = RGB(255, 0, 0)
            wsResults.Cells(r, 6).Font.Bold = True
            logMessages = logMessages & "Row " & r & ": '" & participantName & "' not found"
            If skipReason <> "" Then logMessages = logMessages & " (" & skipReason & ")"
            logMessages = logMessages & vbCrLf
        End If
        GoTo NextRow
        
ApplyResult:
        Select Case UCase(passOrFail)
            Case "PASS"
                targetCell.Value = "Yes"
                If Not targetCell.Comment Is Nothing Then targetCell.Comment.Delete
                targetCell.AddComment "Completed: " & trainingDate
                totalPassed = totalPassed + 1
                wsResults.Range(wsResults.Cells(r, 1), wsResults.Cells(r, 5)).Interior.Color = RGB(198, 239, 206)
            Case "FAIL"
                targetCell.Value = "No"
                If Not targetCell.Comment Is Nothing Then targetCell.Comment.Delete
                targetCell.AddComment "Failed: " & trainingDate
                totalFailed = totalFailed + 1
                wsResults.Range(wsResults.Cells(r, 1), wsResults.Cells(r, 5)).Interior.Color = RGB(255, 199, 206)
            Case "N/A"
                targetCell.Value = "N/A"
                totalNA = totalNA + 1
                wsResults.Range(wsResults.Cells(r, 1), wsResults.Cells(r, 5)).Interior.Color = RGB(255, 235, 156)
        End Select
        Return
        
NextRow:
    Next r
    
    RefreshOnboardingReport
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Dim summary As String
    summary = "Sync Complete!" & vbCrLf & vbCrLf & _
              "  Processed:  " & totalProcessed & vbCrLf & _
              "  Passed:     " & totalPassed & vbCrLf & _
              "  Failed:     " & totalFailed & vbCrLf & _
              "  N/A:        " & totalNA & vbCrLf & _
              "  Not found:  " & totalNotFound & vbCrLf & _
              "  Skipped:    " & totalSkipped & " (already synced)"
    
    If logMessages <> "" Then
        summary = summary & vbCrLf & vbCrLf & "NOTES:" & vbCrLf & logMessages
    End If
    
    MsgBox summary, vbInformation, "Sync Complete"
End Sub

