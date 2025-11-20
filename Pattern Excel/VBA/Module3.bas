Attribute VB_Name = "Module3"
Sub FormatTimesPL2()
Dim ws As Worksheet
Dim lastrow As Long
Dim i As Long
Dim cellValue As String



Set ws = ThisWorkbook.Sheets("PL2 Paste")

lastrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

For i = 7 To lastrow
cellValue = ws.Cells(i, 1).Value

If InStr(1, cellValue, "am", vbTextCompare) > 0 Then
ws.Cells(i, 1).Value = Replace(cellValue, "am", "", 1, -1, vbTextCompare)
cellValue = ws.Cells(i, 1).Value


End If

If InStr(1, cellValue, "pm", vbTextCompare) > 0 Then
ws.Cells(i, 1).Value = Replace(cellValue, "pm", "", 1, -1, vbTextCompare)
cellValue = ws.Cells(i, 1).Value

End If

Next i

For i = 7 To lastrow

cellValue = ws.Cells(i, 1).Value
If InStr(1, cellValue, "(", vbTextCompare) > 0 Then
ws.Cells(i, 1).Value = Replace(cellValue, "(", "   (", 1, -1, vbTextCompare)
cellValue = ws.Cells(i, 1).Value

End If
Next i



End Sub


Sub fixMidnightErrorPL1()
Dim PLWS As Worksheet
Dim lastrow As Long
Dim mainRow As Long
Dim i As Long
Dim timeValue As String
Dim minLastRow As Long
Dim minuteValues As Variant
Dim parts() As String
Dim timeStr As String
Dim startStr As String
Dim endStr As String

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
                End If
                End If
                Next mainRow
            
                
              If UBound(parts) = 1 Then
                    startStr = Trim(parts(0))
                    endStr = Trim(parts(1))
                    
                    ' Handle possible text after end time
                    If InStr(endStr, " ") > 0 Then
                        endStr = Trim(Split(endStr, " ")(0))
                    End If
                    End If
                    
                    MsgBox (startStr)
                    MsgBox (endStr)
                    
                    If startStr > endStr Then
                    

                
                

End Sub


