# VBA-Data-Transform
This is a VBA script I wrote that took input from a spreadsheet and then transformed it into the correct format and outputted a new spreadsheet.

```javascript
Sub makerPro()

    Set ws_Source = Worksheets("Source")
    Set ws_Output = Worksheets("Output")
    Dim rw As Range
    Dim techID As String
    Dim ceOutputRowPosition As Integer

    sInput = InputBox("What week?")
    Dim week As Integer

    If sInput = 1 Then
        week = 2
    ElseIf sInput = 2 Then
        week = 9
    ElseIf sInput = 3 Then
        week = 16
    ElseIf sInput = 4 Then
        week = 23
    ElseIf sInput = 5 Then
        week = 30
    Else
        MsgBox "Bad input"
        Exit Sub
    End If

    ws_Output.Range(Range(Cells(2, 1), Cells(2, 1)), Range(Cells(65535, 1), Cells(65535, 1)).End(xlUp)).Interior.ColorIndex = 0
    For Each rw In ws_Source.Rows

        Set curCell = ws_Source.Cells(rw.Row, 1)
        techID = curCell.Value

        If techID = "" Then
            Exit For
        End If

        ceOutputRowPosition = getPosition(techID)

        If ceOutputRowPosition Then

            Dim outPutCell As Variant
            outPutCell = 3

            ws_Output.Range(Cells(ceOutputRowPosition, 3), Cells(ceOutputRowPosition, 23)).Clear
            ws_Output.Range(Cells(ceOutputRowPosition + 1, 3), Cells(ceOutputRowPosition + 1, 23)).Value = ""

            For col = week To (week + 6)

                Dim min As Variant
                Dim hourminus As Variant
                Set curCellSourceSheet = ws_Source.Cells(rw.Row, col)

                If IsNumeric(curCellSourceSheet) = True And IsEmpty(curCellSourceSheet) = False Then
                
                    If curCellSourceSheet = 15 Then
                        min = "45"
                        hourminus = 1
                    Else
                       min = "00"
                       hourminus = 0
                    End If
            
                    ws_Output.Cells(ceOutputRowPosition, outPutCell).Value = curCellSourceSheet & ":" & "00"
                    outPutCell = outPutCell + 1
                    ws_Output.Cells(ceOutputRowPosition, outPutCell).Value = curCellSourceSheet + 9 - hourminus & ":" & min
                    outPutCell = outPutCell + 1
                    ws_Output.Cells(ceOutputRowPosition, outPutCell).Value = "Working"
                    outPutCell = outPutCell + 1
        
                    If curCellSourceSheet.Interior.Color = RGB(128, 128, 128) Then
                        ws_Output.Cells(ceOutputRowPosition + 1, outPutCell - 3).MergeArea.Value = "All Day Available"
                    End If

                ElseIf curCellSourceSheet.Value = "STD" _
                Or curCellSourceSheet.Value = "T" Or curCellSourceSheet.Value = "TR" _
                Or curCellSourceSheet.Value = "V" Or curCellSourceSheet.Value = "F" _
                Or curCellSourceSheet.Value = "S" Or curCellSourceSheet.Value = "STD" _
                Or curCellSourceSheet.Value = "Sick" _
                And IsEmpty(curCellSourceSheet) = False Then

                    Dim text As String
                    If curCellSourceSheet.Value = "T" Or curCellSourceSheet.Value = "TR" Then
                        text = "Training/Education"
                    ElseIf curCellSourceSheet.Value = "V" Or curCellSourceSheet.Value = "F" Then
                        text = "Vacation"
                    ElseIf curCellSourceSheet.Value = "S" Or curCellSourceSheet.Value = "STD" Or curCellSourceSheet.Value = "Sick" Then
                        text = "Sickness/Medical"
                    End If

                        ws_Output.Cells(ceOutputRowPosition, outPutCell).Value = "08:00"
                        outPutCell = outPutCell + 1
                        ws_Output.Cells(ceOutputRowPosition, outPutCell).Value = "17:00"
                        outPutCell = outPutCell + 1
                        ws_Output.Cells(ceOutputRowPosition, outPutCell).Value = text
                        outPutCell = outPutCell + 1

                Else
                    outPutCell = outPutCell + 3
                    If curCellSourceSheet.Interior.Color = RGB(128, 128, 128) Then
                        ws_Output.Cells(ceOutputRowPosition + 1, outPutCell - 3).MergeArea.Value = "All Day Available"
                    End If
                End If
        
            Next col

            ws_Output.Cells(ceOutputRowPosition, 1).Interior.Color = RGB(0, 255, 0)

        End If

    Next rw

End Sub

Function getPosition(id)

    Set ws = Worksheets("Output")
    Dim rw As Range

    For Each rw In ws.Rows

        Set curCell = ws.Cells(rw.Row, 1)
        techID = curCell.Value

        If curCell.Address = curCell.MergeArea.Cells(1).Address And techID = "" Then
            MsgBox "The end has been reached with out finding a ID " & ws.Cells(rw.Row, 2).Value
            Exit For
        End If

        If InStr(techID, id) > 0 Then
            getPosition = rw.Row
            Exit Function
        End If

    Next rw

    getPosition = False
    MsgBox id & " not found"

End Function

```
