Sub ProcessNewScans()
'
' ProcessNewScans Macro
'
' Keyboard Shortcut: Ctrl+g
'
    Dim Src As String: Src = "<name of source spreadsheet>"
    Dim SrcKey As String: SrcKey = "A" 'The column in which the key field appears.
    Dim SrcRow As Integer: SrcRow = 1 'The row at which data begins.
    Dim SrcQtyKey As String: SrcQtyKey = "C" 'The column in which the quantity appears.
    Dim SrcFlagKey As String: SrcFlagKey = "Z" 'A free column which can be written to for marking rows that have been processed.
    
    Dim Dst As String: Dst = "<name of destination spreadsheet>"
    Dim DstKey As String: DstKey = "A" 'The column in which the key field appears.
    Dim DstRow As Integer: DstRow = 3 'The row at which data begins.
    Dim DstQtyKey As String: DstQtyKey = "L" 'The column in which the quantity is maintained.
    
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
