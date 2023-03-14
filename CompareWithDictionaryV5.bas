Sub CompareWithDictionaryV5()
    On Error GoTo OverflowErrorHandler
    
    ' Declare the variables
    Dim wsPre As Worksheet, wsSAP As Worksheet, wsComp As Worksheet
    Dim countRowPre As Long, countColPre As Long, countRowSAP As Long
    Dim columnValue As Variant
    Dim key As Variant, col1 As Variant, col2 As Variant
    Dim i As Long, x As Long, k As Long
    Dim lHeadColumn As Long
    Dim dict As Object, arrPre As Variant, arrSAP As Variant
    Dim a As Variant, b As Variant
    Dim errorFlag As Boolean
    Dim currentRow As Long
    Dim allCells As Range
    Dim StartTime As Double
    Dim MinutesElapsed As String
    Dim dictArray As Variant
    Dim valueRange As Range
    
    'Turn off screen updating and events
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    errorFlag = False

    StartTime = Timer



    ' Prevent SRC AND SAP
    If ActiveSheet.Name Like "*SRC*" _
        Or ActiveSheet.Name Like "*SAP*" _
        Then
        
        MsgBox "Carefully don't use this worksheet SRC or SAP"
                
        Exit Sub
    End If
    
    ' Define the worksheets
    Set wsPre = Worksheets(ActiveSheet.index - 2)
    Set wsSAP = Worksheets(ActiveSheet.index - 1)
    Set wsComp = ActiveSheet
    Set allCells = ActiveSheet.UsedRange
    
    ' Clear all data at Loaded Sheet
    isClearAll = MsgBox("Do you want to clear all data in this sheet?", vbYesNo + vbQuestion, "Confirm")
    If isClearAll = vbYes Then
        allCells.Clear
    End If
    
    ' Check if worksheet Preload is empty
    If wsPre.Cells(1, 1) = "" Then
            MsgBox "Worksheet Preload file can not be empty"
    ' Check if worksheet SAP is empty
        ElseIf wsSAP.Cells(1, 1) = "" Then
            MsgBox "Worksheet SAP Extraction can not be empty"
    Else
        'Write the title of the first column in the new sheet
        wsComp.Cells(1, 1) = wsPre.Cells(1, 1) & "_Key"

        ' Get the last row and column in wsPre
        countRowPre = wsPre.Cells(Rows.Count, 1).End(xlUp).Row
        countColPre = wsPre.Cells(1, Columns.Count).End(xlToLeft).Column
        
        ' Get the last row and column in wsSAP
        countRowSAP = wsSAP.Cells(Rows.Count, 1).End(xlUp).Row
        countColSAP = wsSAP.Cells(1, Columns.Count).End(xlToLeft).Column 'Use Range from wsPre
        
        ' Copy the values of column A from wsPre to wsComp
        wsComp.Range("A2:A" & countRowPre).Value = wsPre.Range("A2:A" & countRowPre).Value
        
        Set dict = CreateObject("Scripting.Dictionary")
        
        ' SAP make dictionary
        For i = 2 To countRowSAP
            keyValue = wsSAP.Cells(i, 1).Value
            Set customObject = CreateObject("Scripting.Dictionary")
        
            Set valueRange = wsSAP.Range(wsSAP.Cells(i, 1).Address, wsSAP.Cells(i, countColSAP).Address)
            
            currentRow = i
            
            If dict.Exists(keyValue) Then
                MsgBox "Error !!! Comparison is not completly " & vbNewLine & vbNewLine & vbNewLine & "SAP Key has duplicated value: " & keyValue
                Exit Sub
            Else
                dict.Add keyValue, valueRange
            End If
        Next i
        
        'Check Run Time
        Dim MinutesDictionary As String
        MinutesDictionary = Format((Timer - StartTime) / 86400, "hh:mm:ss")
        
        ' Loop through the head columns of wsPre
        For x = 2 To countColPre
            lHeadColumn = wsComp.Cells(1, Columns.Count).End(xlToLeft).Column
            
            ' Add the column headers to wsComp
            wsComp.Cells(1, lHeadColumn + 1) = wsPre.Cells(1, x) & "_SRC"
            wsComp.Cells(1, lHeadColumn + 2) = wsPre.Cells(1, x) & "_SAP"
            wsComp.Cells(1, lHeadColumn + 3) = wsPre.Cells(1, x) & "_COMP"
        Next x
        
        ' Loop through copy the column of wsComp_SRC
        indexColumn = 2
        For x = 2 To (countColPre - 1) * 3
            wsComp.Range(wsComp.Cells(2, x).Address, wsComp.Cells(countRowPre, x).Address).Value = wsPre.Range(wsPre.Cells(2, indexColumn).Address, wsPre.Cells(countRowPre, indexColumn).Address).Value
            
            x = x + 2
            indexColumn = indexColumn + 1
        Next x

        Dim outputArray() As Variant
        Dim NumRows As Long, NumCols As Long
        
        NumRows = wsPre.Cells(Rows.Count, 1).End(xlUp).Row
        NumCols = wsPre.Cells(1, Columns.Count).End(xlToLeft).Column
        
        ReDim outputArray(2 To NumRows, 1 To (NumCols - 1) * 3 + 1)
        
        ' Loop through the rows of wsPre
         For i = 2 To countRowPre
             key = wsPre.Cells(CLng(i), 1).Value
             
             If dict.Exists(key) Then
                indexColumn = 2
                indexSrc = 0
                
                For x = 3 To (countColPre - 1) * 3
                    If x < 0 Then
                        'Comp_SRC is copied
                    Else
                        lHeadColumn = wsComp.Cells(1, Columns.Count).End(xlToLeft).Column
                        indexColumnComp = indexColumn + indexSrc
                        
                        Set a = Nothing
                        Set b = Nothing
                        
                        valueOfA = wsComp.Cells(CLng(i), CLng(indexColumnComp)).Value
                        valueOfB = dict(key).Range(Cells(1, indexColumn).Address).Value
                        
                        If IsError(valueOfA) Then
                            MsgBox "Error!!! On SRC Field" & vbNewLine & vbNewLine & vbNewLine & " Key : " & key
                            
                            Exit Sub
                        ElseIf IsError(valueOfB) Then
                            MsgBox "Error!!! On SAP Field" & vbNewLine & vbNewLine & vbNewLine & " Key : " & key
                            
                            Exit Sub
                        ElseIf valueOfA = "" And valueOfB = "" Then
                            a = ""
                            b = ""
                        ElseIf IsNumeric(valueOfA) And IsNumeric(valueOfB) Then
                            a = CDbl(valueOfA)
                            b = CDbl(valueOfB)
                        Else
                            a = valueOfA
                            b = valueOfB
                        End If
                        
                        'wsComp.Cells(CLng(i), x + 0) = a
                        'wsComp.Cells(CLng(i), x + 0) = IIf(a = "" And b = 0, "", b)
                        'wsComp.Cells(CLng(i), x + 1) = IIf(a = b, "TRUE", "FALSE")
                    
                        outputArray(i, x + 0) = IIf(a = "" And b = 0, "", b)
                        outputArray(i, x + 1) = IIf(a = b, "TRUE", "FALSE")
            
                        x = x + 2
                        indexColumn = indexColumn + 1
                        indexSrc = indexSrc + 2
                    End If
                Next x
                
                dict.Remove key
             Else
                MsgBox "Error !!! Comparison is not completly " & vbNewLine & vbNewLine & vbNewLine & "SRC don't have Key : " & key
                Exit Sub
             End If
         Next i
        
        SetColumnsWithArray outputArray, wsComp, 3, ((NumCols - 1) * 3 + 1), 3
        
        MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")

        
        'Turn on screen updating and events
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        MsgBox "Comparison Complete" & vbNewLine & vbNewLine & vbNewLine & "This code ran in " & MinutesDictionary & " minutes" & " to " & MinutesElapsed & " minutes", vbInformation
    End If

    
    Exit Sub
    
OverflowErrorHandler:
    MsgBox "An Overflow error has occurred. On SAP Line " & currentRow & vbNewLine & vbNewLine & vbNewLine & "Please , Check format value not show ############# , #N/A , REF"
    
End Sub

Function SetColumnsWithArray(outputArray As Variant, wsComp As Worksheet, startCol As Long, endCol As Long, modCol As Long)
    Dim endRow As Long
    Dim columnValues() As Variant
    Dim col As Long
    Dim i As Long
    Dim Rng As Range
    
    Dim chunkSize As Long
    chunkSize = 30000
    
    
    endRow = UBound(outputArray, 1)
    
    For col = startCol To endCol
        ' Get the values from the current column
        ReDim columnValues(1 To endRow) ' Resize the 1D array to hold the column values
        For i = 2 To endRow
            columnValues(i) = outputArray(i, col) ' Get the value from the current column
        Next i
        
        If (col - 1) Mod modCol = 2 Mod modCol Or (col - 1) Mod modCol = 0 Then
            Dim selectedValues() As Variant
            Dim multipleChunkSize As Long
            multipleChunkSize = 1
                
            ' Set the range value to the column values
            For Row = 2 To endRow + 1 Step chunkSize
                Erase selectedValues
                ReDim selectedValues(0 To chunkSize)
                Dim index As Long
                
                If endRow - Row > chunkSize Then
                    index = 0
                    For s = Row To (chunkSize * multipleChunkSize) + 1
                        selectedValues(index) = columnValues(s)
                        
                        index = index + 1
                    Next s
                    
                    Set Rng = wsComp.Range(wsComp.Cells(Row, col), wsComp.Cells(Row + chunkSize, col))
                    Rng.Value = WorksheetFunction.Transpose(selectedValues)
                    
                    multipleChunkSize = multipleChunkSize + 1
                Else
                    index = 0
                    For s = Row To endRow
                        selectedValues(index) = columnValues(s)
                    
                        index = index + 1
                    Next s
                    
                    Set Rng = wsComp.Range(wsComp.Cells(Row, col), wsComp.Cells(endRow, col))
                    Rng.Value = WorksheetFunction.Transpose(selectedValues)
                End If
            Next Row
            
        End If
    Next col
End Function

