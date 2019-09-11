Function fact(a)
    x = 1
    For i = 1 To a
        x = x * i
    Next i
    fact = x
End Function
Sub nopButton()
    Dim i As Integer
    
    Worksheets("ParamNames").Cells.Clear
    Worksheets("ParamNames").Cells(8, 8).Value = "Edit Grey Boxes"
    Worksheets("ParamNames").Cells(15, 8).Value = "Parameter names:"
    Worksheets("ParamNames").Cells(15, 9).Value = "Number of values:"
    Worksheets("ParamNames").Cells(15, 8).Interior.ColorIndex = Worksheets("NumberOfParams").Cells(15, 8).Interior.ColorIndex
    Worksheets("ParamNames").Cells(15, 9).Interior.ColorIndex = Worksheets("NumberOfParams").Cells(15, 8).Interior.ColorIndex
    Worksheets("ParamNames").Cells(8, 8).Font.Bold = True
    Worksheets("ParamNames").Cells(15, 8).Font.Bold = True
    Worksheets("ParamNames").Cells(15, 9).Font.Bold = True

    For i = 1 To Worksheets("NumberOfParams").Cells(15, 9).Value
         Worksheets("ParamNames").Cells(15 + i, 8).Value = "Param" & i
         Worksheets("ParamNames").Cells(15 + i, 8).Interior.ColorIndex = Worksheets("NumberOfParams").Cells(15, 9).Interior.ColorIndex
         Worksheets("ParamNames").Cells(15 + i, 9).Value = 2
         Worksheets("ParamNames").Cells(15 + i, 9).Interior.ColorIndex = Worksheets("NumberOfParams").Cells(15, 9).Interior.ColorIndex
    Next i
    Worksheets("ParamNames").Activate
End Sub

Sub pnButton()
    Dim i As Integer
    Dim j As Integer
    
    Worksheets("ParamValues").Cells.Clear
    Worksheets("ParamValues").Cells(2, 8).Value = "Edit Grey Boxes. Enter Parameter Values."
    Worksheets("ParamValues").Cells(2, 8).Font.Bold = True
    
    For i = 1 To Worksheets("NumberOfParams").Cells(15, 9).Value
         Worksheets("ParamValues").Cells(4 + i, 3).Value = Worksheets("ParamNames").Cells(15 + i, 8)
         Worksheets("ParamValues").Cells(4 + i, 3).Interior.ColorIndex = Worksheets("NumberOfParams").Cells(15, 8).Interior.ColorIndex
    Next i
    
    For i = 1 To Worksheets("NumberOfParams").Cells(15, 9).Value
        For j = 1 To Worksheets("ParamNames").Cells(15 + i, 9).Value
            Worksheets("ParamValues").Cells(4 + i, 3 + j).Value = "P" & i & j
            Worksheets("ParamValues").Cells(4 + i, 3 + j).Interior.ColorIndex = Worksheets("NumberOfParams").Cells(15, 9).Interior.ColorIndex
         Next j
    Next i
    
    Worksheets("ParamValues").Activate
End Sub

Sub pvButton()
    Dim numParam
    Dim i As Integer
    Dim numRows
    numRows = 1
    numParam = Worksheets("NumberOfParams").Cells(15, 9).Value
    Worksheets("Cases").Cells.Clear
    
    For i = 1 To numParam
        numRows = numRows * Worksheets("ParamNames").Cells(15 + i, 9)
    Next i
    
    For paramCounter = 1 To numParam
    
        cellCounter = 0
        divisor = 1
        
        Worksheets("Cases").Cells(4, 3 + paramCounter).Value = Worksheets("ParamNames").Cells(15 + paramCounter, 8)
        Worksheets("Cases").Cells(4, 3 + paramCounter).Interior.ColorIndex = Worksheets("NumberOfParams").Cells(15, 8).Interior.ColorIndex
        
        For divisorCounter = 1 To paramCounter
            divisor = divisor * Worksheets("ParamNames").Cells(15 + divisorCounter, 9)
        Next divisorCounter

        For i = 1 To divisor / Worksheets("ParamNames").Cells(15 + paramCounter, 9)
            For j = 1 To Worksheets("ParamNames").Cells(15 + paramCounter, 9)
                For k = 1 To numRows / divisor
                    cellCounter = cellCounter + 1
                    Worksheets("Cases").Cells(4 + cellCounter, 3 + paramCounter).Value = Worksheets("ParamValues").Cells(4 + paramCounter, 3 + j).Value
                    Worksheets("Cases").Cells(4 + cellCounter, 3 + paramCounter).Interior.ColorIndex = 43
                    Worksheets("Cases").Cells(4 + cellCounter, 3 + paramCounter).Font.Bold = True
                Next k
            Next j
        Next i


    Next paramCounter
    
Worksheets("Cases").Activate
    
End Sub

Sub optimizeButton()
    numRows = 1
    numParam = Worksheets("NumberOfParams").Cells(15, 9).Value
    numComb = fact(numParam) / (fact(2) * fact(numParam - 2))
    ReDim matchFlag(numComb)
    
    For i = 1 To numParam
        numRows = numRows * Worksheets("ParamNames").Cells(15 + i, 9)
    Next i
    
    i = 1
    While i <= numRows
        deleteFlag = 1
        
        For arrayCounter = 0 To numComb - 1
            matchFlag(arrayCounter) = 0
        Next arrayCounter
        
        For j = 1 To numRows
            comb = 0
            If j <> i Then
            
                For m = 1 To numParam - 1
                    For n = m + 1 To numParam
                        If Worksheets("Cases").Cells(4 + i, 3 + m).Value = Worksheets("Cases").Cells(4 + j, 3 + m).Value And Worksheets("Cases").Cells(4 + i, 3 + n).Value = Worksheets("Cases").Cells(4 + j, 3 + n).Value Then
                            matchFlag(comb) = 1
                        End If
                        comb = comb + 1
                    Next n
                Next m
            
            End If
        Next j

        For arrayCounter = 0 To numComb - 1
            If matchFlag(arrayCounter) <> 1 Then
                deleteFlag = 0
            End If
        Next arrayCounter

        If deleteFlag = 1 Then
            Worksheets("Cases").Rows(4 + i).Delete
            i = i - 1
            numRows = numRows - 1
        End If
        i = i + 1
    Wend
End Sub




