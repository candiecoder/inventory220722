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
        Dim ThisKey As String: ThisKey = Year(ThisDate) & "'" & Right("00" & Month(ThisDate), 2) & "'" & Right("00" & Day(ThisDate), 2)
        Dim LastKey As String: LastKey = Sheets(Dst).Cells(DstHdrRow, Asc(DstWeeklyCol) - Asc("A") + 1)
        If Len(LastKey) = 10 Then
            'If a valid date is there, use it as the beginning date for adding new columns.
            If Val(Mid(LastKey, 1, 4)) >= 2023 And Mid(LastKey, 5, 1) = "'" And Val(Mid(LastKey, 6, 2)) > 0 And Mid(LastKey, 8, 1) = "'" And Val(Mid(LastKey, 9, 2)) > 0 Then
                'Convert string date format of "YYYY-MM-DD" to date variable type (so we can do math on it).
                Dim LastDate As Date: LastDate = DateSerial(Val(Mid(LastKey, 1, 4)), Val(Mid(LastKey, 6, 2)), Val(Mid(LastKey, 9, 2)))
                'Keep adding columns until we have enough.
                While LastDate < ThisDate
                    'Increment to the next week.
                    LastDate = DateAdd("d", 7, LastDate)
                    Dim NextKey As String: NextKey = Year(LastDate) & "'" & Right("00" & Month(LastDate), 2) & "'" & Right("00" & Day(LastDate), 2)
                    Sheets(Dst).Range(DstWeeklyCol & DstHdrRow).EntireColumn.Insert
                    Sheets(Dst).Cells(DstHdrRow, Asc(DstWeeklyCol) - Asc("A") + 1) = NextKey
                Wend
            End If
        End If
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
