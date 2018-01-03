Attribute VB_Name = "ScaleHiNData"
Sub ScaleCueSheets()
Attribute ScaleCueSheets.VB_Description = "Macro recorded 2/25/2005 by Alex Wu"
Attribute ScaleCueSheets.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
' Macro recorded 2/25/2005 by Alex Wu
'

'
    
    ChDir "C:\meg\meg_graphs"
    
    Dim strArSubjects(12) As String
    
    'strArSubjects(0) = "ALL"
    strArSubjects(1) = "CJ"
    strArSubjects(2) = "GA"
    strArSubjects(3) = "GN"
    strArSubjects(4) = "HM"
    strArSubjects(5) = "KC"
    strArSubjects(6) = "LB"
    strArSubjects(7) = "ML"
    strArSubjects(8) = "ND"
    strArSubjects(9) = "PS"
    strArSubjects(10) = "VJ"
    strArSubjects(11) = "WA"
    
    'For intSubject = LBound(strArSubjects) To UBound(strArSubjects)
    For intSubject = 1 To UBound(strArSubjects)
    
        strSubject = strArSubjects(intSubject)
        pptfile = "C:\meg\meg_graphs\HiN_" & strSubject & ".xls"
        
        
        Workbooks.Open Filename:=pptfile
        
        
        
        Sheets("CueLeft").Select
        
        Cellvalue = Range("A1").Value
        If (Cellvalue < 10 ^ -12) Then
            
            Sheets("CueLeft").Copy After:=Sheets(3)
            Range("A1").Select
            ActiveCell.FormulaR1C1 = "=CueLeft!RC*10^15"
            Range("A1").Select
            Range(Selection, Selection.End(xlToRight)).Select
            Selection.FillRight
            Range("A1:BO2161").Select
            Selection.FillDown
            Selection.Copy
            Sheets("CueLeft").Select
            Range("A1").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Cells.Select
            Cells.EntireColumn.AutoFit
            Range("A1").Select
            Sheets("CueLeft (2)").Select
            Application.CutCopyMode = False
            ActiveWindow.SelectedSheets.Delete
            
            Sheets("CueRight").Select
            Sheets("CueRight").Copy After:=Sheets(3)
            Range("A1").Select
            ActiveCell.FormulaR1C1 = "=CueRight!RC*10^15"
            Range("A1").Select
            Range(Selection, Selection.End(xlToRight)).Select
            Selection.FillRight
            Range("A1:BO2161").Select
            Selection.FillDown
            Selection.Copy
            Sheets("CueRight").Select
            Range("A1").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Cells.Select
            Cells.EntireColumn.AutoFit
            Range("A1").Select
            Sheets("CueRight (2)").Select
            Application.CutCopyMode = False
            ActiveWindow.SelectedSheets.Delete
        
        End If
        
        ActiveWorkbook.Save
        ActiveWorkbook.Close
    
    Next intSubject

End Sub

Sub RemoveCols()
'
' Macro1 Macro
' Macro recorded 2/25/2005 by Alex Wu
'

'
    
    datapath = "E:\brainstorm_repository\studies\"
    ChDir datapath
    
    Dim strArSubjects(12) As String
    
    'strArSubjects(0) = "ALL"
    strArSubjects(1) = "CJ"
    strArSubjects(2) = "GA"
    strArSubjects(3) = "GN"
    strArSubjects(4) = "HM"
    strArSubjects(5) = "KC"
    strArSubjects(6) = "LB"
    strArSubjects(7) = "ML"
    strArSubjects(8) = "ND"
    strArSubjects(9) = "PS"
    strArSubjects(10) = "VJ"
    strArSubjects(11) = "WA"
    
    'For intSubject = LBound(strArSubjects) To UBound(strArSubjects)
    For intSubject = 1 To UBound(strArSubjects)
    
        strSubject = strArSubjects(intSubject)
        pptfile = datapath & "HiN_" & strSubject & ".xls"
        
        
        Workbooks.Open Filename:=pptfile
        
        Sheets("CueLeft").Select
        Columns("BP:EH").Select
        Selection.Delete Shift:=xlToLeft
        
        Sheets("CueRight").Select
        Columns("BP:EH").Select
        Selection.Delete Shift:=xlToLeft
        
        ActiveWorkbook.Save
        ActiveWorkbook.Close
    
    Next intSubject

End Sub

Sub SummaryStats()
'
' Macro1 Macro
' Macro recorded 2/25/2005 by Alex Wu
'

'
    
    datapath = "E:\brainstorm_repository\studies\"
    ChDir datapath
    
    Dim strArSubjects(12) As String
    
    'strArSubjects(0) = "ALL"
    strArSubjects(1) = "CJ"
    strArSubjects(2) = "GA"
    strArSubjects(3) = "GN"
    strArSubjects(4) = "HM"
    strArSubjects(5) = "KC"
    strArSubjects(6) = "LB"
    strArSubjects(7) = "ML"
    strArSubjects(8) = "ND"
    strArSubjects(9) = "PS"
    strArSubjects(10) = "VJ"
    strArSubjects(11) = "WA"
    
    'For intSubject = LBound(strArSubjects) To UBound(strArSubjects)
    For intSubject = 1 To UBound(strArSubjects)
    
        strSubject = strArSubjects(intSubject)
        pptfile = datapath & "HiN_" & strSubject & ".xls"
        
        Workbooks.Open Filename:=pptfile
        
        Sheets("CueLeft").Select
        ActiveCell.FormulaR1C1 = "=MIN(R[-2163]C:R[-3]C)"
        Range("A2164").Select
        ActiveCell.FormulaR1C1 = "=MIN(R[-2163]C:R[-3]C)"
        Range("A2164:A2165").Select
        Selection.FillDown
        Range("A2165").Select
        ActiveCell.FormulaR1C1 = "=MAX(R[-2164]C:R[-4]C)"
        Range("A2164:BO2165").Select
        Selection.FillRight
        Range("B2164").Select
        Selection.End(xlToRight).Select
        Range("BQ2165").Select
        
        Sheets("CueRight").Select
        ActiveCell.FormulaR1C1 = "=MIN(R[-2163]C:R[-3]C)"
        Range("A2164").Select
        ActiveCell.FormulaR1C1 = "=MIN(R[-2163]C:R[-3]C)"
        Range("A2164:A2165").Select
        Selection.FillDown
        Range("A2165").Select
        ActiveCell.FormulaR1C1 = "=MAX(R[-2164]C:R[-4]C)"
        Range("A2164:BO2165").Select
        Selection.FillRight
        Range("B2164").Select
        Selection.End(xlToRight).Select
        Range("BQ2165").Select
        
        ActiveWorkbook.Save
        ActiveWorkbook.Close
    
    Next intSubject

End Sub

