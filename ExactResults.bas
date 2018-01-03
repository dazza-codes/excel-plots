Attribute VB_Name = "ExactResults_Module"

Sub exact_1_load_results_sum_table()
Attribute exact_1_load_results_sum_table.VB_Description = "Macro recorded 03/04/00 by Darren Weber"
Attribute exact_1_load_results_sum_table.VB_ProcData.VB_Invoke_Func = " \n14"
'
' exact_1_load_results_sum_table Macro
' Macro coded 03/04/00 by Darren Weber
'

'

    Dim path, exactSheet As String
    
    path = "D:\My Documents\THESIS\results\erp results\divergence analysis\"
    
    group = Array("cont", "ptsd")
    
    cond = Array("sa", "wm", "ea", "dt")
    
    
    ' Exact test sheet sorting parameters
    Dim Row, Col, NoContent, _
        TSUMcritCol, TSUMobservCol, _
        TABSUMcritCol, TABSUMobservCol, _
        TMAXcritCol, TMAXobservCol As Integer
    
    TSUMcritCol = 4
    TSUMobservCol = 5
    
    TABSUMcritCol = 7
    TABSUMobservCol = 8
    
    TMAXcritCol = 10
    TMAXobservCol = 11
        
    ' Tmax significant values scanning parameters
    Dim Time, Value, NextValue, Sig1, Sig2, NextSig1, NextSig2 As Double, _
        StartCol1, StartCol2, break As Integer, _
        GotStart1, GotStart2, GotTmax1, GotTmax2, GotEnd1, GotEnd2, Tail As String
    
    For g = 0 To UBound(group)
        For cd = 0 To UBound(cond)
            
                      
            exactSheet = (cond(cd) & "_" & group(g) & "_div")
            
            Workbooks.OpenText FileName:=(path & exactSheet & ".sum") _
                , Origin:=xlWindows, StartRow:=1, DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, _
                Tab:=True, Semicolon:=False, Comma:=False, Space:=True, _
                Other:=True, OtherChar:=",", FieldInfo:= _
                Array( _
                Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), _
                Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), _
                Array(9, 1) _
                )
            ActiveWorkbook.SaveAs FileName:=(path & exactSheet & ".xls") _
                , FileFormat:=xlNormal, Password:="", WriteResPassword:="", _
                ReadOnlyRecommended:=False, CreateBackup:=False
            ActiveWorkbook.Close
            
            Workbooks.Open FileName:=(path & exactSheet & ".xls")
            Sheets(exactSheet).Select
            
            Row = 1
            Rows(Row).Select
            Selection.Insert Shift:=xlDown
            Rows(Row + 1).Select
            Selection.Insert Shift:=xlDown
            
            ' Insert Column Headings
        
            Cells(Row, 4) = "TSUM"
            Cells(Row, 7) = "ABSUM"
            Cells(Row, 10) = "TMAX"
            Col = 4
            While (Col <= 10)
                Cells(Row + 1, Col) = "Crit"
                Cells(Row + 1, Col + 1) = "Observ"
                Cells(Row + 1, Col + 2) = "p ="
                Col = Col + 3
            Wend
            Cells(Row + 1, Col) = "start"
            Cells(Row + 1, Col + 1) = "peak"
            Cells(Row + 1, Col + 2) = "end"
            Cells(Row + 1, Col + 3) = "tail"
            Cells(Row + 1, Col + 4) = "sig"
            
            Row = 3
            Col = 1
            NoContent = 0
            
            While (NoContent < 20)
                
                Cells(Row, 1).Select
                
                If (Cells(Row, 1) = "") Then
                    
                    Rows(Row).Select
                    Selection.Delete Shift:=xlUp
                    NoContent = NoContent + 1
                
                ElseIf ((InStr(1, Cells(Row, 2), "L", vbTextCompare) = 1) Xor _
                        (InStr(1, Cells(Row, 2), "R", vbTextCompare) = 1)) Then
                
                    Cells(Row, 1).Delete Shift:=xlToLeft
                    Range(Cells(Row, 1), Cells(Row + 1, 3)).Select
                    Selection.FillDown
                    
                    NoContent = 0
                    Row = Row + 2
                    
                ElseIf (StrComp(Cells(Row, 2), "observed", vbTextCompare) = 0) Then
                    
                    If (StrComp(Cells(Row, 1), "TSUM", vbTextCompare) = 0) Then
                        PasteCol = TSUMobservCol
                    ElseIf (StrComp(Cells(Row, 1), "ABSUM", vbTextCompare) = 0) Then
                        PasteCol = TABSUMobservCol
                    ElseIf (StrComp(Cells(Row, 1), "TMAX", vbTextCompare) = 0) Then
                        PasteCol = TMAXobservCol
                    End If
                    
                    ' Copy&Paste observed
                    Cells(Row, 5).Select
                    Selection.Copy
                    Cells(Row - 2, PasteCol).Select
                    ActiveSheet.PASTE
                    Range(Cells(Row - 2, PasteCol), Cells(Row - 1, PasteCol)).Select
                    Selection.FillDown
                    
                    ' Copy&Paste 1-tail observed significance
                    Cells(Row, 8).Select
                    Selection.Copy
                    Cells(Row - 2, PasteCol + 1).Select
                    ActiveSheet.PASTE
                    ' Multiply 1-tail sig. value by 2
                    Cells(Row - 1, PasteCol + 1) = 2 * Cells(Row - 2, PasteCol + 1)
                    
                    Rows(Row).Select
                    Selection.Delete Shift:=xlUp
                    NoContent = 0
                    
                ElseIf (StrComp(Cells(Row, 2), "critical", vbTextCompare) = 0) Then
                    
                    If (StrComp(Cells(Row, 1), "TSUM", vbTextCompare) = 0) Then
                        PasteCol = TSUMcritCol
                    ElseIf (StrComp(Cells(Row, 1), "ABSUM", vbTextCompare) = 0) Then
                        PasteCol = TABSUMcritCol
                    ElseIf (StrComp(Cells(Row, 1), "TMAX", vbTextCompare) = 0) Then
                        PasteCol = TMAXcritCol
                    End If
                    
                    ' Cut&Paste 1-tail critical value @ 0.05
                    Cells(Row + 8, 3).Select
                    Selection.Copy
                    Cells(Row - 2, PasteCol).Select
                    ActiveSheet.PASTE
                    
                    ' Cut&Paste 2-tail critical value @ 0.05
                    Cells(Row + 8, 4).Select
                    Selection.Copy
                    Cells(Row - 1, PasteCol).Select
                    ActiveSheet.PASTE
                    
                    For r = 0 To 9
                        Rows(Row).Select
                        Selection.Delete Shift:=xlUp
                    Next r
                    
                    NoContent = 0
                    
                ElseIf (StrComp(Cells(Row, 1), "Significant", vbTextCompare) = 0) Then
                
                    '#################################
                    'Scan Tmax values for start,max,end of significance
                    
                    ' Delete "significance" row and next empty row
                    Rows(Row).Select
                    Selection.Delete Shift:=xlUp
                    Rows(Row).Select
                    Selection.Delete Shift:=xlUp
                    
                    StartCol1 = 13
                    StartCol2 = 13
                    GotStart1 = "n"
                    GotStart2 = "n"
                    GotTmax1 = "n"
                    GotTmax2 = "n"
                    GotEnd1 = "n"
                    GotEnd2 = "n"

                    If (Cells(Row, 2) = "") Then
                        ' No significant Tmax values
                        break = 1
                    Else
                        Time = Cells(Row, 2)
                        Value = Cells(Row, 3)
                        Sig1 = Cells(Row, 6)
                        Sig2 = Cells(Row, 11)
                        If (Cells(Row, 8) Like "a<b*") Then
                            Tail = "a<b"
                        ElseIf (Cells(Row, 8) Like "a>b*") Then
                            Tail = "a>b"
                        End If
                        
                        If (Sig1 <= 0.05) Then
                            Cells(Row - 2, StartCol1) = Time
                            Cells(Row - 2, StartCol1 + 3) = Tail
                            GotStart1 = "y"
                        End If
                        If (Sig2 <= 0.05) Then
                            Cells(Row - 1, StartCol2) = Time
                            Cells(Row - 1, StartCol2 + 3) = Tail
                            GotStart2 = "y"
                        End If
                        
                        Rows(Row).Select
                        Selection.Delete Shift:=xlUp
                        
                        break = 0
                    End If
                    
                    While (break < 1)
                        
                        If (GotStart1 = "n" And Sig1 <= 0.05) Then
                            Cells(Row - 2, StartCol1) = Time
                            Cells(Row - 2, StartCol1 + 3) = Tail
                            GotStart1 = "y"
                        End If
                        If (GotStart2 = "n" And Sig2 <= 0.05) Then
                            Cells(Row - 1, StartCol2) = Time
                            Cells(Row - 1, StartCol2 + 3) = Tail
                            GotStart2 = "y"
                        End If
                        
                        NextTime = Cells(Row, 2)
                        NextValue = Cells(Row, 3)
                        NextSig1 = Cells(Row, 6)
                        NextSig2 = Cells(Row, 11)
                        
                        ' Check for end of all significance values
                        If (NextValue = "") Then
                            
                            ' Finish scanning Tmax values
                            break = 1
                            
                            ' Set end values, unless set already
                            If (GotEnd1 = "n" And Sig1 <= 0.05) Then
                                Cells(Row - 2, StartCol1 + 2) = Time
                            End If
                            If (GotEnd2 = "n" And Sig2 <= 0.05) Then
                                Cells(Row - 1, StartCol2 + 2) = Time
                            End If
                        
                        ' Check for a break in 1-tailed significance values
                        ElseIf (Sig1 <= 0.05 And NextSig1 > 0.05) Then
                            
                            Cells(Row - 2, StartCol1 + 2) = Time
                            
                            GotEnd1 = "n"
                            StartCol1 = StartCol1 + 5
                            GotStart1 = "n"
                            GotTmax1 = "n"
                            
                        ' Check for a break in 2-tailed significance values
                        ElseIf (Sig2 <= 0.05 And NextSig2 > 0.05) Then
                            
                            Cells(Row - 1, StartCol2 + 2) = Time
                            
                            GotEnd2 = "n"
                            StartCol2 = StartCol2 + 5
                            GotStart2 = "n"
                            GotTmax2 = "n"
                            
                        ' Test for skip in time points
                        ElseIf ((Time + 2.5) <> NextTime) Then
                            
                            ' define end points, unless already set
                            If (GotEnd1 = "n" And Sig1 <= 0.05) Then
                                Cells(Row - 2, StartCol1 + 2) = Time
                            End If
                            If (GotEnd2 = "n" And Sig2 <= 0.05) Then
                                Cells(Row - 1, StartCol2 + 2) = Time
                            End If
                            
                            StartCol1 = StartCol1 + 5
                            GotStart1 = "n"
                            GotTmax1 = "n"
                            GotEnd1 = "n"
                            
                            StartCol2 = StartCol2 + 5
                            GotStart2 = "n"
                            GotTmax2 = "n"
                            GotEnd2 = "n"
                        
                        ' Check for Tmax values
                        ElseIf ((Value >= 0 And Value > NextValue) Or (Value < 0 And Value < NextValue)) Then

                            If (GotTmax1 = "n" And Sig1 <= 0.05) Then
                                Cells(Row - 2, StartCol1 + 1) = Time
                                Cells(Row - 2, StartCol1 + 4) = Sig1
                                GotTmax1 = "y"
                            End If
                            If (GotTmax2 = "n" And Sig2 <= 0.05) Then
                                Cells(Row - 1, StartCol2 + 1) = Time
                                Cells(Row - 1, StartCol2 + 4) = Sig2
                                GotTmax2 = "y"
                            End If
                        
                        End If
                        
                        If (break = 0) Then
                            
                            ' Continue searching
                            Time = Cells(Row, 2)
                            Value = Cells(Row, 3)
                            Sig1 = Cells(Row, 6)
                            Sig2 = Cells(Row, 11)
                            If (Cells(Row, 8) Like "a<b*") Then
                                Tail = "a<b"
                            ElseIf (Cells(Row, 8) Like "a>b*") Then
                                Tail = "a>b"
                            End If
                            
                            ' Delete Scanned Row
                            Rows(Row).Select
                            Selection.Delete Shift:=xlUp
                        
                        End If
                    Wend
                    
                    NoContent = 0
                    
                ElseIf (StrComp(Cells(Row, 1), "Minimum", vbTextCompare) = 0) Then
                    
                    Rows(Row).Select
                    Selection.Delete Shift:=xlUp
                    NoContent = 0
                    
                End If
                
            Wend
            
            Columns.Select
            Selection.NumberFormat = "0.00"
                        
            Cells(1, 1).Select
            ActiveWorkbook.Save
            ActiveWorkbook.Close
        
        Next cd
    Next g
End Sub
Sub exact_2_split_1and2tail_results_table()
'
' exact_2_split_1and2tail_results_table Macro
' Macro coded 10/04/00 by Darren Weber
'
'

    Dim path, exactSheet As String
    path = "D:\My Documents\THESIS\results\erp results\divergence analysis\"
    
    group = Array("cont", "ptsd")
    
    cond = Array("sa", "wm", "ea", "dt")
    
    ' Exact test sheet sorting parameters
    Dim Row, Col, NoContent, sigEmpty, sigCol As Integer
    
    For g = 0 To UBound(group)
        For cd = 0 To UBound(cond)

            exactSheet = (cond(cd) & "_" & group(g) & "_div")
            
            Workbooks.Open FileName:=(path & exactSheet & ".xls")
            Sheets(exactSheet).Select
            
            Sheets(exactSheet).Copy Before:=Sheets(1)
            Sheets(exactSheet).Copy Before:=Sheets(2)
            
            Sheets(exactSheet & " (2)").Name = (exactSheet & " 1-tailed")
            Sheets(exactSheet & " (3)").Name = (exactSheet & " 2-tailed")
            
            tailSheets = Array(exactSheet & " 1-tailed", exactSheet & " 2-tailed")
            
            For s = 0 To UBound(tailSheets)
            
                Sheets(tailSheets(s)).Select
                
                NoContent = 0
                sigEmpty = 0
                sigCol = 13
                Row = 3
                
                If (InStr(1, tailSheets(s), "2-tailed", vbTextCompare) > 0) Then
                    Rows(3).Select
                    Selection.Delete Shift:=xlUp
                End If
                
                While (NoContent < 1)
                    
                    If (StrComp(Cells(Row, 1), "", vbTextCompare) = 0) Then
                        NoContent = 1
                    End If
                    
                    While (sigEmpty < 10)
                        If (StrComp(Cells(Row, sigCol), "", vbTextCompare) = 0) Then
                            Range(Cells(Row, sigCol), Cells(Row, sigCol + 4)).Select
                            Selection.Delete Shift:=xlToLeft
                            sigEmpty = sigEmpty + 1
                        Else
                            sigCol = sigCol + 5
                        End If
                    Wend
                    sigEmpty = 0
                    sigCol = 13

                    Rows(Row + 1).Select
                    Selection.Delete Shift:=xlUp
                    
                    Row = Row + 1
                    
                Wend
            
                Cells(1, 1).Select
            
            Next s
            

            ActiveWorkbook.Save
            ActiveWorkbook.Close
            
        Next cd
    Next g
End Sub
Sub exact_3_only_sig_results_table()
'
' exact_3_only_sig_results_table Macro
' Macro coded 15/04/00 by Darren Weber
'
'

    Dim path, exactSheet As String
    path = "D:\My Documents\THESIS\results\erp results\divergence analysis\"
    
    group = Array("cont", "ptsd")
   
    cond = Array("sa", "wm", "ea", "dt")
    
    
    ' Exact test sheet sorting parameters
    Dim Row, Col, Sig, Start, Finish, NoContent As Integer
    
    For g = 0 To UBound(group)
        For cd = 0 To UBound(cond)

            exactSheet = (cond(cd) & "_" & group(g) & "_div")
            
            Workbooks.Open FileName:=(path & exactSheet & ".xls")
            Sheets(exactSheet).Select
            
            Sheets(exactSheet & " 2-tailed").Copy Before:=Sheets(exactSheet & " 1-tailed")
            Sheets(exactSheet & " 2-tailed (2)").Name = (exactSheet & " 2-tailed sig")
            
            Sheets(exactSheet & " 1-tailed").Copy Before:=Sheets(exactSheet & " 2-tailed sig")
            Sheets(exactSheet & " 1-tailed (2)").Name = (exactSheet & " 1-tailed sig")
                                    
            sigSheets = Array(exactSheet & " 1-tailed sig", exactSheet & " 2-tailed sig")
                                    
            For s = 0 To UBound(sigSheets)
                
                NoContent = 0
                Row = 3
                
                Sheets(sigSheets(s)).Select
                
                While (NoContent < 1)
                    
                    If (StrComp(Cells(Row, 1), "", vbTextCompare) = 0) Then
                        NoContent = 1
                    Else
                        
                        If (StrComp(Cells(Row, 18), "", vbTextCompare) <> 0) Then
                            Rows(Row).Select
                            Selection.Copy
                            Selection.Insert Shift:=xlDown
                            
                            Range(Cells(Row, 18), Cells(Row, 60)).Select
                            Selection.Delete Shift:=xlToLeft
                            
                            Range(Cells(Row + 1, 13), Cells(Row + 1, 17)).Select
                            Selection.Delete Shift:=xlToLeft
                                                        
                        End If
                                                
                        Sig = Cells(Row, 12)
                        Start = Cells(Row, 13)
                        Finish = Cells(Row, 15)
                        
                        If (Sig > 0.06) Then
                            Rows(Row).Select
                            Selection.Delete Shift:=xlUp
                        ElseIf (Finish <= (Start + 10)) Then
                            Rows(Row).Select
                            Selection.Font.Italic = True
                            Row = Row + 1
                        Else
                            Row = Row + 1
                        End If
                        
                    End If
                Wend
                
                ' Code to sort the table by signif time points rather than regions
                '
                ' Columns("J:K").Select
                ' Selection.Sort Key1:=Range(""), Order1:=xlAscending, Key2:=Range("K2") _
                    , Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
                    False, Orientation:=xlTopToBottom
                
                Cells.Select
                Cells.EntireColumn.AutoFit
                Cells(3, 4).Select
                ActiveWindow.FreezePanes = True
                
            Next s
            
            ActiveWorkbook.Save
            ActiveWorkbook.Close
            
        Next cd
    Next g
End Sub

Sub exact_4_cleanup_region_names()
'
' exact_4_cleanup_region_names Macro
' Macro coded 16/01/01 by Darren Weber
'
'

    Dim path, exactSheet As String
    path = "D:\My Documents\THESIS\results\erp results\divergence analysis\"
    
    group = Array("cont", "ptsd")
   
    cond = Array("sa", "wm", "ea", "dt")
    
    For g = 0 To UBound(group)
        For cd = 0 To UBound(cond)

            exactSheet = (cond(cd) & "_" & group(g) & "_div")
            
            Workbooks.Open FileName:=(path & exactSheet & ".xls")

            Dim sh As Worksheet

            For Each sh In ActiveWorkbook.Worksheets
            
                sh.Select
                
                If (StrComp(Cells(4, 3), "", vbTextCompare) = 0) Then
                    
                    Columns(1).Select
                    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
                        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
                        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
                        :="_", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1))
                End If
                    
                Cells.Select
                Cells.EntireColumn.AutoFit
                Cells(3, 4).Select
                ActiveWindow.FreezePanes = True
                            
            Next sh
            
            ActiveWorkbook.Save
            ActiveWorkbook.Close
            
        Next cd
    Next g

End Sub
