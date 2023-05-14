Attribute VB_Name = "Module2"
Sub ProcessNewScans()
'
' ProcessNewScans Macro
'
' Keyboard Shortcut: Ctrl+g
'
    Dim Src As String: Src = "Stocking Activity"
    Dim SrcKey As String: SrcKey = "A" 'The column in which the key field appears.
    Dim SrcRow As Integer: SrcRow = 1 'The row at which data begins.
    Dim SrcQtyKey As String: SrcQtyKey = "C" 'The column in which the quantity appears.
    Dim SrcDateKey As String: SrcDateKey = "D" 'The column in which the date appears.
    Dim SrcFlagKey As String: SrcFlagKey = "Z" 'A free column which can be written to for marking rows that have been processed.
    
    Dim Dst As String: Dst = "Stockroom"
    Dim DstKey As String: DstKey = "A" 'The column in which the key field appears.
    Dim DstRow As Integer: DstRow = 3 'The row at which data begins.
    Dim DstQtyKey As String: DstQtyKey = "L" 'The column in which the quantity is maintained.
    'For weekly column:
    Dim DstWeeklyCol As String: DstWeeklyCol = "AC" 'The first weekly column.
    Dim DstHdrRow As Integer: DstHdrRow = 2 'The row where the header belongs.
    
    Dim Ans As VbMsgBoxResult
    Dim I As Integer, J As Integer, M As Integer, N As Integer
    
    'Computed values.
    Dim SrcIdx As Integer: SrcIdx = Sheets(Src).Columns(SrcKey).Column
    Dim SrcQtyIdx As Integer: SrcQtyIdx = Sheets(Src).Columns(SrcQtyKey).Column
    Dim SrcDateIdx As Integer: SrcDateIdx = Sheets(Src).Columns(SrcDateKey).Column
    Dim SrcFlagIdx As Integer: SrcFlagIdx = Sheets(Src).Columns(SrcFlagKey).Column
    Dim DstIdx As Integer: DstIdx = Sheets(Dst).Columns(DstKey).Column
    Dim DstQtyIdx As Integer: DstQtyIdx = Sheets(Dst).Columns(DstQtyKey).Column
    Dim DstWeeklyIdx As Integer: DstWeeklyIdx = Sheets(Dst).Columns(DstWeeklyCol).Column

    'Count new rows and count how many of them are found in the destination sheet.
    I = SrcRow
    M = 0: N = 0
    While Sheets(Src).Cells(I, SrcIdx) <> ""
        If Sheets(Src).Cells(I, SrcFlagIdx) = "" Then
            N = N + 1 'Count new rows.
            J = DstRow
            While Sheets(Dst).Cells(J, DstIdx) <> ""
                If Sheets(Dst).Cells(J, DstIdx) _
                 = Sheets(Src).Cells(I, SrcIdx) Then
                    M = M + 1 'Count how many rows are found in the destination sheet.
                    J = 32766
                End If
                J = J + 1
            Wend
        End If
        I = I + 1
    Wend
    
    'Prompt the user based on the counts determined above.
    If N = 0 Then
        Ans = MsgBox("No new rows found in """ & Src & """.", vbExclamation & vbOK): Ans = vbCancel
    ElseIf M = 0 Then
        Ans = MsgBox("No new matching rows from """ & Src & """ found in """ & Dst & """.", vbExclamation & vbOK): Ans = vbCancel
    Else
        Ans = MsgBox("Import " & M & " of " & N & " new rows from """ & Src & """ to """ & Dst & """?", vbQuestion & vbOKCancel)
    End If
    If Ans <> vbOK Then
        Exit Sub
    End If
        
    'Add weekly column(s) if needed.
    If Sheets(Dst).Cells(DstHdrRow, DstWeeklyIdx) <> "" Then
        Dim ThisDate As Date: ThisDate = Now
        Dim LastDate As Date
        On Error Resume Next
        LastDate = Sheets(Dst).Range(DstWeeklyCol & DstHdrRow).Value
        If Err.Number = 0 Then
            'Keep adding columns until we have enough.
            While LastDate < ThisDate
                'Increment to the next week.
                LastDate = DateAdd("d", 7, LastDate)
                Sheets(Dst).Range(DstWeeklyCol & DstHdrRow).EntireColumn.Insert
                Sheets(Dst).Range(DstWeeklyCol & DstHdrRow).Value = LastDate
                Sheets(Dst).Range(DstWeeklyCol & DstHdrRow).EntireColumn.Insert
                Sheets(Dst).Range(DstWeeklyCol & DstHdrRow).Value = LastDate
            Wend
        End If
        On Error GoTo 0
    End If
    
    'Process the new scans from the source spreadsheet, updating the destination sheet.
    I = SrcRow
    M = 0: N = 0
    While Sheets(Src).Cells(I, SrcIdx) <> ""
        If Sheets(Src).Cells(I, SrcFlagIdx) = "" Then
            N = N + 1
            J = DstRow
            While Sheets(Dst).Cells(J, DstIdx) <> ""
                If Sheets(Dst).Cells(J, DstIdx) _
                 = Sheets(Src).Cells(I, SrcIdx) Then
                 
                    'Update the quantity in the destination sheet.
                    Sheets(Dst).Cells(J, DstQtyIdx) = _
                    Sheets(Dst).Cells(J, DstQtyIdx) + _
                    Sheets(Src).Cells(I, SrcQtyIdx)
                    
                    'Get the date of the transaction.
                    Dim TxDate As Date
                    TxDate = Sheets(Src).Range(SrcDateKey & I).Value
                    
                    'Find the correct weekly summary column.
                    Dim W As Integer
                    Dim StartDate As Date
                    For W = 0 To 100 Step 2
                        If Sheets(Dst).Cells(DstHdrRow, DstWeeklyIdx + W) = "" Then Exit For
                        On Error Resume Next
                        StartDate = Sheets(Dst).Cells(DstHdrRow, DstWeeklyIdx + W).Value
                        If Err.Number = 0 Then
                            If TxDate >= StartDate Then
                            
                                'Update the quantity in the weekly column.
                                If Sheets(Src).Cells(I, SrcQtyIdx) < 0 Then
                                    Sheets(Dst).Cells(J, DstWeeklyIdx + W) = _
                                    Val(Sheets(Dst).Cells(J, DstWeeklyIdx + W)) + _
                                    Sheets(Src).Cells(I, SrcQtyIdx)
                                Else
                                    Sheets(Dst).Cells(J, DstWeeklyIdx + 1 + W) = _
                                    Val(Sheets(Dst).Cells(J, DstWeeklyIdx + 1 + W)) + _
                                    Sheets(Src).Cells(I, SrcQtyIdx)
                                End If
                                Exit For
                            End If
                        End If
                        On Error GoTo 0
                    Next
                    
                    'Mark the source row as done.
                    Sheets(Src).Cells(I, SrcFlagIdx) = "Done"
                    
                    M = M + 1
                    J = 32766
                End If
                J = J + 1
            Wend
        End If
        I = I + 1
    Wend
    MsgBox "Done.", vbInformation & vbOKOnly
End Sub
