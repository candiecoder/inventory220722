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
    Dim DstWeeklyCol As String: DstWeeklyCol = "N" 'The first weekly column.
    Dim DstHdrRow As Integer: DstHdrRow = 2 'The row where the header belongs.
    
    Dim Ans As VbMsgBoxResult
    Dim I As Integer, J As Integer, M As Integer, N As Integer
    
    'Count new rows and count how many of them are found in the destination sheet.
    I = SrcRow
    M = 0: N = 0
    While Sheets(Src).Cells(I, Asc(SrcKey) - Asc("A") + 1) <> ""
        If Sheets(Src).Cells(I, Asc(SrcFlagKey) - Asc("A") + 1) = "" Then
            N = N + 1 'Count new rows.
            J = DstRow
            While Sheets(Dst).Cells(J, Asc(DstKey) - Asc("A") + 1) <> ""
                If Sheets(Dst).Cells(J, Asc(DstKey) - Asc("A") + 1) _
                 = Sheets(Src).Cells(I, Asc(SrcKey) - Asc("A") + 1) Then
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
    If Sheets(Dst).Cells(DstHdrRow, Asc(DstWeeklyCol) - Asc("A") + 1) <> "" Then
        Dim ThisDate As Date: ThisDate = Now
        Dim LastDate As Date
        On Error Resume Next
        LastDate = Sheets(Dst).Range(DstWeeklyCol & DstHdrRow).Value2
        If Err.Number = 0 Then
            'Keep adding columns until we have enough.
            While LastDate < ThisDate
                'Increment to the next week.
                LastDate = DateAdd("d", 7, LastDate)
                Sheets(Dst).Range(DstWeeklyCol & DstHdrRow).EntireColumn.Insert
                Sheets(Dst).Range(DstWeeklyCol & DstHdrRow).Value = LastDate
            Wend
        End If
        On Error GoTo 0
    End If
    
    'Process the new scans from the source spreadsheet, updating the destination sheet.
    I = SrcRow
    M = 0: N = 0
    While Sheets(Src).Cells(I, Asc(SrcKey) - Asc("A") + 1) <> ""
        If Sheets(Src).Cells(I, Asc(SrcFlagKey) - Asc("A") + 1) = "" Then
            N = N + 1
            J = DstRow
            While Sheets(Dst).Cells(J, Asc(DstKey) - Asc("A") + 1) <> ""
                If Sheets(Dst).Cells(J, Asc(DstKey) - Asc("A") + 1) _
                 = Sheets(Src).Cells(I, Asc(SrcKey) - Asc("A") + 1) Then
                 
                    'Update the quantity in the destination sheet.
                    Sheets(Dst).Cells(J, Asc(DstQtyKey) - Asc("A") + 1) = _
                    Sheets(Dst).Cells(J, Asc(DstQtyKey) - Asc("A") + 1) + _
                    Sheets(Src).Cells(I, Asc(SrcQtyKey) - Asc("A") + 1)
                    
                    'Get the date of the transaction.
                    Dim TxDate As Date
                    TxDate = Sheets(Src).Range(SrcDateKey & I).Value2
                    
                    'Find the correct weekly summary column.
                    Dim W As Integer
                    Dim StartDate As Date
                    For W = 1 To 100
                        If Sheets(Dst).Cells(DstHdrRow, Asc(DstWeeklyCol) - Asc("A") + W) = "" Then Exit For
                        On Error Resume Next
                        StartDate = Sheets(Dst).Cells(DstHdrRow, Asc(DstWeeklyCol) - Asc("A") + W).Value2
                        If Err.Number = 0 Then
                            If TxDate >= StartDate Then
                            
                                'Update the quantity in the weekly column.
                                Sheets(Dst).Cells(J, Asc(DstWeeklyCol) - Asc("A") + W) = _
                                Val(Sheets(Dst).Cells(J, Asc(DstWeeklyCol) - Asc("A") + W)) + _
                                Sheets(Src).Cells(I, Asc(SrcQtyKey) - Asc("A") + 1)
                                
                                Exit For
                            End If
                        End If
                        On Error GoTo 0
                    Next
                    
                    'Mark the source row as done.
                    Sheets(Src).Cells(I, Asc(SrcFlagKey) - Asc("A") + 1) = "Done"
                    
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
