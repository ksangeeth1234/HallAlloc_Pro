Attribute VB_Name = "Module1"
Option Explicit

' ============================================================
' RunHallSearch
' Reads Day (B7), Time (E7) and Month (B8) from the
' "Hall Availability Search" sheet, checks every lecture hall
' against the "Class Schedule Analysis" sheet, and writes
' Free / Occupied results into the results table (rows 14-41).
'
' The Month check looks at each class's Start Date / End Date
' (columns P and Q on Class Schedule Analysis) and only counts
' the class as occupying the hall if the selected month falls
' within that date range. Choose "(Any Month)" in B8 to skip
' this check (matches the previous Day+Time-only behaviour).
' ============================================================
Sub RunHallSearch()

    Dim wsSearch As Worksheet
    Dim wsData As Worksheet
    Dim searchDay As String
    Dim searchTime As Variant
    Dim searchMinutes As Long
    Dim monthCellVal As Variant
    Dim useMonthFilter As Boolean
    Dim monthStart As Date, monthEnd As Date
    Dim lastDataRow As Long
    Dim r As Long, j As Long
    Dim hallName As String
    Dim found As Boolean
    Dim dStart As Long, dEnd As Long
    Dim dDay As String, dHall As String, dStatus As String
    Dim dStartDate As Variant, dEndDate As Variant
    Dim monthPasses As Boolean

    On Error GoTo ErrHandler

    Set wsSearch = ThisWorkbook.Sheets("Hall Availability Search")
    Set wsData = ThisWorkbook.Sheets("Class Schedule Analysis")

    searchDay = Trim(wsSearch.Range("B7").Value)
    searchTime = wsSearch.Range("E7").Value

    If Not IsDate(searchTime) And Not IsNumeric(searchTime) Then
        MsgBox "Please enter a valid time in E7 (e.g. 9:00 AM).", vbExclamation, "Invalid time"
        Exit Sub
    End If
    searchMinutes = Hour(CDate(searchTime)) * 60 + Minute(CDate(searchTime))

    ' ---- Month filter setup ----
    monthCellVal = wsSearch.Range("B8").Value
    useMonthFilter = False
    If Not (VarType(monthCellVal) = vbString) Then
        ' a real date was selected -> filter by that month
        useMonthFilter = True
        monthStart = DateSerial(Year(CDate(monthCellVal)), Month(CDate(monthCellVal)), 1)
        monthEnd = DateSerial(Year(monthStart), Month(monthStart) + 1, 1) - 1
    ElseIf IsDate(monthCellVal) Then
        useMonthFilter = True
        monthStart = DateSerial(Year(CDate(monthCellVal)), Month(CDate(monthCellVal)), 1)
        monthEnd = DateSerial(Year(monthStart), Month(monthStart) + 1, 1) - 1
    End If
    ' If B8 = "(Any Month)" (text), useMonthFilter stays False

    lastDataRow = wsData.Cells(wsData.Rows.Count, "B").End(xlUp).Row

    Application.ScreenUpdating = False

    Dim firstResultRow As Long, lastResultRow As Long
    firstResultRow = 14
    lastResultRow = wsSearch.Cells(wsSearch.Rows.Count, "A").End(xlUp).Row
    If lastResultRow < firstResultRow Then lastResultRow = firstResultRow

    For r = firstResultRow To lastResultRow
        hallName = wsSearch.Cells(r, 1).Value
        If hallName = "" Then Exit For

        found = False
        Dim occBy As String, occMode As String, occTime As String
        occBy = "": occMode = "": occTime = ""

        For j = 3 To lastDataRow
            dDay = wsData.Cells(j, 2).Value            ' column B: Weekday
            dHall = wsData.Cells(j, 10).Value           ' column J: Lecture Hall
            dStatus = wsData.Cells(j, 19).Value         ' column S: Class Status

            If dDay = searchDay And dHall = hallName And _
               (dStatus = "Ongoing" Or dStatus = "Upcoming") Then

                dStart = wsData.Cells(j, 7).Value        ' column G: Start (min)
                dEnd = wsData.Cells(j, 8).Value          ' column H: End (min)

                If searchMinutes >= dStart And searchMinutes < dEnd Then

                    monthPasses = True
                    If useMonthFilter Then
                        dStartDate = wsData.Cells(j, 16).Value  ' column P: Start Date
                        dEndDate = wsData.Cells(j, 17).Value    ' column Q: End Date
                        monthPasses = False
                        If IsDate(dStartDate) And IsDate(dEndDate) Then
                            If CDate(dStartDate) <= monthEnd And CDate(dEndDate) >= monthStart Then
                                monthPasses = True
                            End If
                        End If
                    End If

                    If monthPasses Then
                        found = True
                        occBy = wsData.Cells(j, 12).Value    ' column L: Program / Batch
                        occMode = wsData.Cells(j, 13).Value  ' column M: Mode
                        occTime = wsData.Cells(j, 5).Value & " - " & wsData.Cells(j, 6).Value
                        Exit For
                    End If
                End If
            End If
        Next j

        If found Then
            wsSearch.Cells(r, 3).Value = "Occupied"
            wsSearch.Cells(r, 4).Value = occBy
            wsSearch.Cells(r, 5).Value = occMode
            wsSearch.Cells(r, 6).Value = occTime
        Else
            wsSearch.Cells(r, 3).Value = "Free"
            wsSearch.Cells(r, 4).Value = ""
            wsSearch.Cells(r, 5).Value = ""
            wsSearch.Cells(r, 6).Value = ""
        End If
    Next r

    Dim monthLabel As String
    If useMonthFilter Then
        monthLabel = Format(monthStart, "mmm yyyy")
    Else
        monthLabel = "Any Month"
    End If

    wsSearch.Range("B11").Value = "Last searched: " & searchDay & " at " & _
        Format(CDate(searchTime), "h:mm AM/PM") & ", " & monthLabel & _
        " (run " & Format(Now, "dd-mmm-yyyy hh:mm") & ")"

    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Something went wrong: " & Err.Description, vbCritical, "Hall Search Error"

End Sub

' ============================================================
' Optional: auto-run the search whenever Day, Time or Month is
' changed. To enable this, open the VBA editor (Alt+F11),
' double-click the "Hall Availability Search" sheet under
' Microsoft Excel Objects, and paste the code below into THAT
' sheet's module (not this one).
' ============================================================
'Private Sub Worksheet_Change(ByVal Target As Range)
'    If Not Intersect(Target, Me.Range("B7:B7,E7:E7,B8:B8")) Is Nothing Then
'        RunHallSearch
'    End If
'End Sub
