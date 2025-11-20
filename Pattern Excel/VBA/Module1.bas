Attribute VB_Name = "Module1"

Sub Pl1All()
ClearRowEStart
StartEndTimePL1
CalcEmptPL1
CalcTestPL1
ExportPL1CopyToNotepad
End Sub



Sub ClearRowEStart()
Dim PLBD As Worksheet
Set PLBD = ThisWorkbook.Sheets("PL1 Broken Down")
PLBD.Range("E6:E5000").ClearContents
End Sub


Sub StartEndTimePL1()
    Dim wsBrokenDown As Worksheet
    Dim wsPaste As Worksheet
    Dim lastrow As Long
    Dim minuteValues As Variant
    Dim minuteDbl() As Double
    Dim b3Time As Date, b4Time As Date, c3Time As Date, c4Time As Date
    Dim currentTime As Date
    Dim currentTimeDbl As Double
    Dim tolerance As Double
    Dim i As Long, j As Long
    Dim n As Long
    Dim endTimeAdjusted As Date

    ' tolerance (fraction of day). Keep small — adjust if needed.
    tolerance = 0.00000001

    Set wsBrokenDown = ThisWorkbook.Worksheets("PL1 Broken Down")
    Set wsPaste = ThisWorkbook.Worksheets("PL1 Paste")

    ' find last row with times in column B (B6 downward)
    lastrow = wsBrokenDown.Cells(wsBrokenDown.Rows.Count, "B").End(xlUp).Row
    If lastrow < 6 Then Exit Sub

    ' load minute values into variant array
    minuteValues = wsBrokenDown.Range("B6:B" & lastrow).Value
    n = UBound(minuteValues, 1)
    ReDim minuteDbl(1 To n)

    ' normalize minute values to time fraction (0..1)
    For j = 1 To n
        If IsNumeric(minuteValues(j, 1)) Then
            minuteDbl(j) = CDbl(minuteValues(j, 1)) - Int(CDbl(minuteValues(j, 1)))
        Else
            minuteDbl(j) = -1 ' invalid marker
        End If
    Next j

    ' Step 1: set E6:E1445 to purple text exactly (per your request)
    wsBrokenDown.Range("E6:E1445").Value = "8000ff"

    ' read times from PL1 Paste (no heavy validation per your instruction)
    On Error Resume Next
    b3Time = wsPaste.Range("B3").Value
    b4Time = wsPaste.Range("B4").Value
    c3Time = wsPaste.Range("C3").Value
    c4Time = wsPaste.Range("C4").Value
    On Error GoTo 0

    ' Helper: mark a start->end range (handles midnight crossing by adding 1 day when end < start)
    ' We'll reuse the exact matching style you used but with normalization + tolerance
    If b3Time > 0 And b4Time > 0 Then
        currentTime = b3Time
        endTimeAdjusted = b4Time
        If endTimeAdjusted < currentTime Then endTimeAdjusted = endTimeAdjusted + 1 ' cross-midnight guard
        Do While currentTime <= endTimeAdjusted + (tolerance * 10)
            currentTimeDbl = CDbl(currentTime) - Int(CDbl(currentTime)) ' normalized fraction
            For j = 1 To n
                If minuteDbl(j) >= 0 Then
                    If Abs(minuteDbl(j) - currentTimeDbl) <= tolerance Then
                        wsBrokenDown.Range("E" & (5 + j)).Value = "59bf6b"
                        Exit For
                    End If
                End If
            Next j
            currentTime = DateAdd("n", 1, currentTime)
        Loop
    End If

    ' Second range (C3 -> C4) — allows crossing midnight
    If c3Time > 0 And c4Time > 0 Then
        currentTime = c3Time
        endTimeAdjusted = c4Time
        If endTimeAdjusted < currentTime Then endTimeAdjusted = endTimeAdjusted + 1 ' next day
        Do While currentTime <= endTimeAdjusted + (tolerance * 10)
            currentTimeDbl = CDbl(currentTime) - Int(CDbl(currentTime)) ' normalized fraction
            For j = 1 To n
                If minuteDbl(j) >= 0 Then
                    If Abs(minuteDbl(j) - currentTimeDbl) <= tolerance Then
                        wsBrokenDown.Range("E" & (5 + j)).Value = "59bf6b"
                        Exit For
                    End If
                End If
            Next j
            currentTime = DateAdd("n", 1, currentTime)
        Loop
    End If

    ' done
    ' (optional) MsgBox "Done"
End Sub






Sub CalcEmptPL1()
    Dim PLWS As Worksheet
    Dim PLBD As Worksheet
    Dim mainRow As Long
    Dim minLastRow As Long
    Dim timeStr As String
    Dim parts() As String
    Dim startStr As String
    Dim endStr As String
    Dim startTime As Date
    Dim endTime As Date
    Dim currentTime As Date
    Dim foundRow As Variant
    Dim minuteValues As Variant
    Dim i As Long
    Dim currentTimeDbl As Double
    Dim tolerance As Double
    Dim matchFound As Boolean
    
    tolerance = 0.0000007 ' About 0.06 seconds
    
    Set PLWS = ThisWorkbook.Sheets("PL1 Paste")
    Set PLBD = ThisWorkbook.Sheets("PL1 Broken Down")

    minLastRow = PLBD.Cells(PLBD.Rows.Count, "B").End(xlUp).Row
    minuteValues = PLBD.Range("B6:B" & minLastRow).Value
    
    For mainRow = 7 To 200
        If PLWS.Range("B" & mainRow).Value = True Then
            timeStr = Trim(PLWS.Range("D" & mainRow).Value)
            If Len(timeStr) > 0 Then
                ' Normalize dash types (sometimes Excel pastes with special en dash)
                timeStr = Replace(timeStr, "–", "-")
                parts = Split(timeStr, "-")
                
                If UBound(parts) = 1 Then
                    startStr = Trim(parts(0))
                    endStr = Trim(parts(1))
                    
                    ' Handle possible text after end time
                    If InStr(endStr, " ") > 0 Then
                        endStr = Trim(Split(endStr, " ")(0))
                    End If
                    
                    On Error Resume Next
                    startTime = timeValue(startStr)
                    endTime = timeValue(endStr)
                    On Error GoTo 0
                    
                      ' Handle overnight range (e.g., 23:50–00:30)
                     If endTime < startTime Then
                         endTime = endTime + 1
                     End If
                    
                    If startTime >= 0 And endTime > startTime Then
                        currentTime = startTime
                        Do While currentTime <= endTime
                            currentTimeDbl = CDbl(currentTime)
                            matchFound = False
                            
                            ' Try to match with tolerance
                            For i = 1 To UBound(minuteValues, 1)
                                If IsNumeric(minuteValues(i, 1)) Then
                                    If Abs(minuteValues(i, 1) - currentTimeDbl) < tolerance Then
                                        PLBD.Range("E" & (5 + i)).Value = "ff6633"
                                        matchFound = True
                                        Exit For
                                    End If
                                End If
                            Next i
                            
                            If Not matchFound Then
                                Debug.Print "No match for " & Format(currentTime, "hh:mm") & _
                                            " (" & currentTimeDbl & ") from row " & mainRow
                            End If
                            
                            currentTime = DateAdd("n", 1, currentTime)
                        Loop
                    Else
                        Debug.Print "Invalid time range in row " & mainRow & ": " & timeStr
                    End If
                End If
            End If
       End If
    Next mainRow
End Sub


Sub CalcTestPL1()
    Dim PLWS As Worksheet
    Dim PLBD As Worksheet
    Dim mainRow As Long
    Dim minLastRow As Long
    Dim timeStr As String
    Dim parts() As String
    Dim startStr As String
    Dim endStr As String
    Dim startTime As Date
    Dim endTime As Date
    Dim currentTime As Date
    Dim foundRow As Variant
    Dim minuteValues As Variant
    Dim i As Long
    Dim currentTimeDbl As Double
    Dim tolerance As Double
    Dim matchFound As Boolean
    
    tolerance = 0.0000007 ' About 0.06 seconds
    
    Set PLWS = ThisWorkbook.Sheets("PL1 Paste")
    Set PLBD = ThisWorkbook.Sheets("PL1 Broken Down")

    minLastRow = PLBD.Cells(PLBD.Rows.Count, "B").End(xlUp).Row
    minuteValues = PLBD.Range("B6:B" & minLastRow).Value
    
    For mainRow = 7 To 200
        If PLWS.Range("C" & mainRow).Value = True Then
            timeStr = Trim(PLWS.Range("D" & mainRow).Value)
            If Len(timeStr) > 0 Then
                ' Normalize dash types (sometimes Excel pastes with special en dash)
                timeStr = Replace(timeStr, "–", "-")
                parts = Split(timeStr, "-")
                
                If UBound(parts) = 1 Then
                    startStr = Trim(parts(0))
                    endStr = Trim(parts(1))
                    
                    ' Handle possible text after end time
                    If InStr(endStr, " ") > 0 Then
                        endStr = Trim(Split(endStr, " ")(0))
                    End If
                    
                    On Error Resume Next
                    startTime = timeValue(startStr)
                    endTime = timeValue(endStr)
                    On Error GoTo 0
                    
                    ' Handle overnight range (e.g., 23:50–00:30)
                    If endTime < startTime Then
                        endTime = endTime + 1
                    End If
                    
                    If startTime >= 0 And endTime > startTime Then
                        currentTime = startTime
                        Do While currentTime <= endTime
                            currentTimeDbl = CDbl(currentTime)
                            matchFound = False
                            
                            ' Try to match with tolerance
                            For i = 1 To UBound(minuteValues, 1)
                                If IsNumeric(minuteValues(i, 1)) Then
                                    If Abs(minuteValues(i, 1) - currentTimeDbl) < tolerance Then
                                        PLBD.Range("E" & (5 + i)).Value = "0099ff"

                                        matchFound = True
                                        Exit For
                                    End If
                                End If
                            Next i
                            
                            If Not matchFound Then
                                Debug.Print "No match for " & Format(currentTime, "hh:mm") & _
                                            " (" & currentTimeDbl & ") from row " & mainRow
                            End If
                            
                            currentTime = DateAdd("n", 1, currentTime)
                        Loop
                    Else
                        Debug.Print "Invalid time range in row " & mainRow & ": " & timeStr
                    End If
                End If
            End If
        End If
    Next mainRow
End Sub


Sub ExportPL1CopyToNotepad()
    Dim ws As Worksheet
    Dim desktopPath As String
    Dim FilePath As String
    Dim lastrow As Long, LastCol As Long
    Dim r As Long, c As Long
    Dim lineText As String
    Dim fso As Object, ts As Object
    Dim safeName As String

    On Error GoTo ErrHandler

    ' --- sheet ---
    Set ws = ThisWorkbook.Sheets("PL1 COPY THIS")
    If ws Is Nothing Then
        MsgBox "Sheet 'PL1 COPY' not found.", vbExclamation
        Exit Sub
    End If

    ' --- desktop path (more reliable than Environ) ---
    desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"

    ' --- safe filename (remove illegal filename chars just in case) ---
    safeName = "PL1 COPY"
    safeName = Replace(safeName, "/", "-")
    safeName = Replace(safeName, "\", "-")
    safeName = Replace(safeName, ":", "-")
    safeName = Replace(safeName, "*", "-")
    safeName = Replace(safeName, "?", "")
    safeName = Replace(safeName, """", "")
    safeName = Replace(safeName, "<", "-")
    safeName = Replace(safeName, ">", "-")
    safeName = Replace(safeName, "|", "-")
    FilePath = desktopPath & safeName & ".txt"

    ' --- safely determine used range (handles empty sheet) ---
    With ws
        If Application.WorksheetFunction.CountA(.Cells) = 0 Then
            lastrow = 1: LastCol = 1
        Else
            lastrow = .Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
            LastCol = .Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
        End If
    End With

    ' --- FileSystemObject to create file (handles permissions better than Open) ---
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' If file exists try delete; if locked, append timestamp instead
    If fso.FileExists(FilePath) Then
        On Error Resume Next
        fso.DeleteFile FilePath, True
        On Error GoTo ErrHandler
        If fso.FileExists(FilePath) Then
            FilePath = desktopPath & safeName & " " & Format(Now, "yyyymmdd_hhnnss") & ".txt"
        End If
    End If

    Set ts = fso.CreateTextFile(FilePath, True, False) ' overwrite=True, unicode=False (ASCII)

    ' --- write rows (tabs between cells) ---
    For r = 1 To lastrow
        lineText = ""
        For c = 1 To LastCol
            lineText = lineText & ws.Cells(r, c).Text
            If c < LastCol Then lineText = lineText & vbTab
        Next c
        ts.WriteLine lineText
    Next r

    ts.Close

    MsgBox "Export complete:" & vbCrLf & FilePath, vbInformation

    ' Optional: open Notepad with the file
    On Error Resume Next
    Shell "notepad.exe """ & FilePath & """", vbNormalFocus
    On Error GoTo 0

    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "File path: " & FilePath, vbCritical
End Sub

