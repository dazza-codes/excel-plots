Attribute VB_Name = "ERP_ParsePeaks"

Sub peaks_all()

    Dim path As String, _
        filenames As String
    
    path = "F:\data_emse\ptsdpet\scd14hz\"
    filenames = "*_n80.peaks"
    
    Call peaks_parse_files(path, filenames)
    
End Sub
Sub peaks_parse_files(path As String, _
                      filenames As String, _
                      Optional subfolders As Boolean = False)
    
    Set fs = Application.FileSearch
    
    With fs
        .NewSearch
        .LookIn = path
        '.SearchSubFolders = subfolders
        .FileName = filenames & ".xls"
        
        If .Execute > 0 Then
            Message = "Do you want to replace all """ & filenames & ".xls"" ?"
            Reply = MsgBox(Message, vbOKCancel, "Replace Files")
            If Reply = vbOK Then
                Kill (path & filenames & ".xls")
            End If
        End If
    End With
    
    With fs
        .NewSearch
        .LookIn = path
        '.SearchSubFolders = subfolders
        .FileName = filenames
        
        If .Execute > 0 Then
            
            'MsgBox "There were " & .FoundFiles.Count & " file(s) found."
            For i = 1 To .FoundFiles.Count
                
                Workbooks.OpenText FileName:=.FoundFiles(i), _
                    Origin:=xlWindows, StartRow:=10, DataType:=xlFixedWidth, FieldInfo:= _
                    Array(Array(0, 1), Array(8, 1), Array(35, 1), Array(52, 1), Array(64, 1), Array(73, 1), _
                    Array(83, 1))
                Columns("A:G").Select
                Columns("A:G").EntireColumn.AutoFit
                Range("A1").Select
                ActiveWorkbook.SaveAs FileName:=.FoundFiles(i), _
                FileFormat:=xlNormal, Password:="", WriteResPassword:="", _
                ReadOnlyRecommended:=False, CreateBackup:=False
                ActiveWorkbook.Close
                ' MsgBox .FoundFiles(i)
            Next i
        Else
            MsgBox "There were no files found."
        End If
    End With
    
End Sub

Sub peaks_reorganise_files()

'   Rearrange peak data for each peak file

    Set fs = Application.FileSearch
    With fs
    
        .LookIn = "D:\My Documents\THESIS\results\erp results\PEAKS"
        .SearchSubFolders = True
        .FileName = "*.peaks.xls"
        
        If .Execute > 0 Then
            'MsgBox "There are " & .FoundFiles.Count & " file(s) to process."
            For i = 1 To .FoundFiles.Count
                
                Workbooks.Open .FoundFiles(i)
                
                Columns("B:B").Select
                Selection.Delete Shift:=xlToLeft
                
                'Test for "REGION" headings in first row and remove first 2 rows
                If (StrComp(Cells(1, 1), "REGION", vbTextCompare) = 0) Then
                    Rows(1).Select
                    Selection.Delete Shift:=xlUp
                    If (StrComp(Cells(1, 1), "", vbTextCompare) = 0) Then
                        Rows(1).Select
                        Selection.Delete Shift:=xlUp
                    End If
                End If
                
                Row = 1
                data = 1
                While (data > 0)
                
                    Cells(Row, 1).Select
                
                    'Copy electrode title down, if necessary
                    If (StrComp(Cells(Row, 1), "", vbTextCompare) = 0) Then
                        Cells(Row - 1, 1).Select
                        Selection.Copy
                        Cells(Row, 1).Select
                        ActiveSheet.PASTE
                    End If
                    
                    'Test for "NA" in positive row and adjust negative data
                    If (StrComp(Cells(Row, 4), "NA", vbTextCompare) = 0) Then
                    
                        ' Copy negative electrode to current row
                        Cells(Row + 1, 3).Select
                        Selection.Copy
                        Cells(Row, 3).Select
                        ActiveSheet.PASTE
                        
                        ' Copy negative values to current row
                        Range(Cells(Row + 1, 5), Cells(Row + 1, 6)).Select
                        Selection.Copy
                        Cells(Row, 4).Select
                        ActiveSheet.PASTE
                        
                        ' Delete next row (now empty)
                        Row = Row + 1
                        Rows(Row).Select
                        Selection.Delete Shift:=xlUp
                        
                    'Test for "No POS" row and remove negative row
                    ElseIf (StrComp(Cells(Row, 4), "No POS", vbTextCompare) = 0) Then
                        
                        ' Delete next negative row (should be "NA")
                        Row = Row + 1
                        Rows(Row).Select
                        Selection.Delete Shift:=xlUp
                        
                    'Test for "NA" in negative row/cell and remove row
                    ElseIf (StrComp(Cells(Row, 5), "NA", vbTextCompare) = 0) Then
                    
                            Rows(Row).Select
                            Selection.Delete Shift:=xlUp
                            
                    'Test for empty cell in positive row and remove it
                    ElseIf (StrComp(Cells(Row, 5), "", vbTextCompare) = 0) Then
                        
                            Cells(Row, 5).Select
                            Selection.Delete Shift:=xlToLeft
                            Row = Row + 1
                            
                    Else
                            Row = Row + 1
                    End If
                
                    'Check for more data (row 2 has values every 2nd row)
                    If (StrComp(Cells(Row, 2), "", vbTextCompare) = 0) And _
                       (StrComp(Cells(Row + 1, 2), "", vbTextCompare) = 0) Then
                        
                        Rows(Row).Select
                        Selection.Delete Shift:=xlUp
                        Rows(Row).Select
                        Selection.Delete Shift:=xlUp
                        
                        data = 0
                    Else
                        data = 1
                    End If
                
                Wend
                
                Cells(1, 1).Select
                ActiveWorkbook.Save
                ActiveWorkbook.Close
            
            Next i
        Else
            MsgBox "There were no files found."
        End If
    End With
    
End Sub

Sub peaks_organise_sw()
'
' peaks_organise_sw Macro
' Macro coded 14/11/00 by Darren Weber
'

'   Reorganise peak data for each slow wave (sw) peak file

    Dim parsedata, data, Row, X, NWindows As Integer
    
    NWindows = 5
    
    Set fs = Application.FileSearch
    With fs
        .LookIn = "D:\My Documents\RESEARCH\twin study\peaks\sw"
        '.SearchSubFolders = True
        .FileName = "*.peaks.xls"
        
        If .Execute > 0 Then
            'MsgBox "There are " & .FoundFiles.Count & " file(s) to process."
            For i = 1 To .FoundFiles.Count
                
                Workbooks.Open .FoundFiles(i)
                
                'Split slow wave values to separate columns
                Row = 1
                parsedata = 0
                While (Row < 20 And parsedata = 0)
                    If (StrComp(Cells(Row, 4), "", vbTextCompare) = 0) Then
                        parsedata = 0
                        Row = Row + 1
                    Else
                        parsedata = 1
                    End If
                Wend
                If (parsedata > 0) Then
                    Columns(4).Select
                    Selection.TextToColumns Destination:=Range("D1"), DataType:=xlDelimited, _
                        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
                        Semicolon:=False, Comma:=False, Space:=True, Other:=True, OtherChar:="(", _
                        FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), _
                        Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 9))
                End If
                
                Cells.Select
                Selection.Columns.AutoFit
                Cells(1, 1).Select
                
                Cells.Replace What:=")", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
                
                'Shift columns of sw data into rows
                data = 0
                Row = 1
                While (data < 1)
                
                    If (StrComp(Cells(Row, 1), "", vbTextCompare) = 0) Then
                        data = 1
                    Else
                        X = 1
                        While (X <= NWindows)
                            Rows(Row).Select
                            Selection.Copy
                            Rows(Row + 1).Select
                            Selection.Insert Shift:=xlDown
                            
                            Row = Row + 1
                            Range(Cells(Row, 3), Cells(Row, 4)).Select
                            Selection.Delete Shift:=xlToLeft
                            
                            X = X + 1
                        Wend
                        
                        Row = Row + 1
                        
                    End If
                Wend
                
                Range(Cells(1, 5), Cells(2000, 200)).Select
                Selection.Clear
                
                Cells(1, 1).Select
                ActiveWorkbook.Save
                ActiveWorkbook.Close
                
            Next i
        Else
            MsgBox "There were no files found."
        End If
    End With

End Sub

Sub peaks_transpose()
'
' peaks_transpose Macro
' Macro recorded 25/10/00 by Darren Weber
'

'   Transpose peak data for each peak file

    Set fs = Application.FileSearch
    With fs
        .LookIn = "D:\My Documents\THESIS\results\erp results\PEAKS"
        '.SearchSubFolders = True
        .FileName = "*.peaks.xls"
        
        If .Execute > 0 Then
            'MsgBox "There are " & .FoundFiles.Count & " file(s) to process."
            For i = 1 To .FoundFiles.Count
                
                Workbooks.Open .FoundFiles(i)
                
                Range(Cells(1, 1), Cells(250, 5)).Select
                Selection.Copy
                
                'Paste (transposed)
                Cells(1, 6).Select
                Selection.PasteSpecial _
                    PASTE:=xlAll, _
                    Operation:=xlNone, _
                    SkipBlanks:=False, _
                    Transpose:=True
                
                'Remove copied area, shift cells left
                Range(Cells(1, 1), Cells(250, 5)).Select
                Selection.Delete Shift:=xlToLeft
                
                Cells.Select
                Selection.Columns.AutoFit
                Cells(1, 1).Select
                
                ActiveWorkbook.Save
                ActiveWorkbook.Close
            
            Next i
        Else
            MsgBox "There were no files found."
        End If
    End With
    
End Sub

Sub ptsd_peaks_compose_all_files()
'
' ptsd_peaks_compose_all_files Macro
' Macro recorded 25/10/00 by Darren Weber
'

'   Gather all peak data for all peak files

    path = "D:\My Documents\THESIS\results\erp results\PEAKS"

    'Define and open all peak files
    cont_elec = path & "\cont_elec.xls"
    cont_amp = path & "\cont_amp.xls"
    cont_lat = path & "\cont_lat.xls"
    ptsd_elec = path & "\ptsd_elec.xls"
    ptsd_amp = path & "\ptsd_amp.xls"
    ptsd_lat = path & "\ptsd_lat.xls"
    
    'Initialise row counters for placement of peak data
    cont_elec_row = 1
    cont_amp_row = 1
    cont_lat_row = 1
    ptsd_elec_row = 1
    ptsd_amp_row = 1
    ptsd_lat_row = 1
    
    Dim foundfile As String
    
    Set fs = Application.FileSearch
    With fs
        .LookIn = path
        '.SearchSubFolders = True
        .FileName = "*.peaks.xls"
        
        If .Execute > 0 Then
        
            Workbooks.Add
            ActiveWorkbook.SaveAs FileName:=cont_elec, AddToMru:=True
            GoSub ADDSHEETS
            Workbooks.Add
            ActiveWorkbook.SaveAs FileName:=cont_amp, AddToMru:=True
            GoSub ADDSHEETS
            Workbooks.Add
            ActiveWorkbook.SaveAs FileName:=cont_lat, AddToMru:=True
            GoSub ADDSHEETS
            Workbooks.Add
            ActiveWorkbook.SaveAs FileName:=ptsd_elec, AddToMru:=True
            GoSub ADDSHEETS
            Workbooks.Add
            ActiveWorkbook.SaveAs FileName:=ptsd_amp, AddToMru:=True
            GoSub ADDSHEETS
            Workbooks.Add
            ActiveWorkbook.SaveAs FileName:=ptsd_lat, AddToMru:=True
            GoSub ADDSHEETS
            
            'MsgBox "There are " & .FoundFiles.Count & " file(s) to process."
            
            For i = 1 To .FoundFiles.Count
                
                Workbooks.Open .FoundFiles(i)
                
                foundfile = right(.FoundFiles(i), 20)
                group = Mid(foundfile, 1, 1)
                cond = Mid(foundfile, 5, 2)
                ' for dif waves use: foundfile = Right(.FoundFiles(i), 22)
                'MsgBox "file = " & foundfile & "group = " & group & "cond = " & cond
                
                
                'Select Electrode row (#3)
                Rows(3).Select
                Selection.Copy
                
                'Paste electrode data into group file
                If (InStr(1, group, "c", vbTextCompare) > 0) Then
                
                    Windows("cont_elec.xls").Activate
                    
                    If (InStr(1, cond, "ea", vbTextCompare) > 0) Then
                        Sheets("ea").Select
                        PASTE_ROW = cont_elec_row
                        GoSub PASTE_ROWS
                    ElseIf (InStr(1, cond, "dt", vbTextCompare) > 0) Then
                        Sheets("dt").Select
                        PASTE_ROW = cont_elec_row
                        GoSub PASTE_ROWS
                    ElseIf (InStr(1, cond, "sa", vbTextCompare) > 0) Then
                        Sheets("sa").Select
                        PASTE_ROW = cont_elec_row
                        GoSub PASTE_ROWS
                    ElseIf (InStr(1, cond, "wm", vbTextCompare) > 0) Then
                        Sheets("wm").Select
                        PASTE_ROW = cont_elec_row
                        GoSub PASTE_ROWS
                        cont_elec_row = cont_elec_row + 1
                    End If
                    
                    ActiveWorkbook.Save
                
                ElseIf (InStr(1, group, "p", vbTextCompare) > 0) Then
                
                    Windows("ptsd_elec.xls").Activate
                    
                    If (InStr(1, cond, "ea", vbTextCompare) > 0) Then
                        Sheets("ea").Select
                        PASTE_ROW = ptsd_elec_row
                        GoSub PASTE_ROWS
                    ElseIf (InStr(1, cond, "dt", vbTextCompare) > 0) Then
                        Sheets("dt").Select
                        PASTE_ROW = ptsd_elec_row
                        GoSub PASTE_ROWS
                    ElseIf (InStr(1, cond, "sa", vbTextCompare) > 0) Then
                        Sheets("sa").Select
                        PASTE_ROW = ptsd_elec_row
                        GoSub PASTE_ROWS
                    ElseIf (InStr(1, cond, "wm", vbTextCompare) > 0) Then
                        Sheets("wm").Select
                        PASTE_ROW = ptsd_elec_row
                        GoSub PASTE_ROWS
                        ptsd_elec_row = ptsd_elec_row + 1
                    End If
                    
                    ActiveWorkbook.Save
                
                End If
                
                
                
                'Select Amplitude row (#4)
                Windows(foundfile).Activate
                Rows(4).Select
                Selection.Copy
                
                'Paste amp data into group file
                If (InStr(1, group, "c", vbTextCompare) > 0) Then
                
                    Windows("cont_amp.xls").Activate
                    
                    If (InStr(1, cond, "ea", vbTextCompare) > 0) Then
                        Sheets("ea").Select
                        PASTE_ROW = cont_amp_row
                        GoSub PASTE_ROWS
                    ElseIf (InStr(1, cond, "dt", vbTextCompare) > 0) Then
                        Sheets("dt").Select
                        PASTE_ROW = cont_amp_row
                        GoSub PASTE_ROWS
                    ElseIf (InStr(1, cond, "sa", vbTextCompare) > 0) Then
                        Sheets("sa").Select
                        PASTE_ROW = cont_amp_row
                        GoSub PASTE_ROWS
                    ElseIf (InStr(1, cond, "wm", vbTextCompare) > 0) Then
                        Sheets("wm").Select
                        PASTE_ROW = cont_amp_row
                        GoSub PASTE_ROWS
                        cont_amp_row = cont_amp_row + 1
                    End If
                    
                    ActiveWorkbook.Save
                
                ElseIf (InStr(1, group, "p", vbTextCompare) > 0) Then
                
                    Windows("ptsd_amp.xls").Activate
                    
                    If (InStr(1, cond, "ea", vbTextCompare) > 0) Then
                        Sheets("ea").Select
                        PASTE_ROW = ptsd_amp_row
                        GoSub PASTE_ROWS
                    ElseIf (InStr(1, cond, "dt", vbTextCompare) > 0) Then
                        Sheets("dt").Select
                        PASTE_ROW = ptsd_amp_row
                        GoSub PASTE_ROWS
                    ElseIf (InStr(1, cond, "sa", vbTextCompare) > 0) Then
                        Sheets("sa").Select
                        PASTE_ROW = ptsd_amp_row
                        GoSub PASTE_ROWS
                    ElseIf (InStr(1, cond, "wm", vbTextCompare) > 0) Then
                        Sheets("wm").Select
                        PASTE_ROW = ptsd_amp_row
                        GoSub PASTE_ROWS
                        ptsd_amp_row = ptsd_amp_row + 1
                    End If
                    
                    ActiveWorkbook.Save
                
                End If
                
                
                'Select latency row (#5)
                Windows(foundfile).Activate
                Rows(5).Select
                Selection.Copy
                
                'Paste lat data into correct condition file
                If (InStr(1, group, "c", vbTextCompare) > 0) Then
                
                    Windows("cont_lat.xls").Activate
                    
                    If (InStr(1, cond, "ea", vbTextCompare) > 0) Then
                        Sheets("ea").Select
                        PASTE_ROW = cont_lat_row
                        GoSub PASTE_ROWS
                    ElseIf (InStr(1, cond, "dt", vbTextCompare) > 0) Then
                        Sheets("dt").Select
                        PASTE_ROW = cont_lat_row
                        GoSub PASTE_ROWS
                    ElseIf (InStr(1, cond, "sa", vbTextCompare) > 0) Then
                        Sheets("sa").Select
                        PASTE_ROW = cont_lat_row
                        GoSub PASTE_ROWS
                    ElseIf (InStr(1, cond, "wm", vbTextCompare) > 0) Then
                        Sheets("wm").Select
                        PASTE_ROW = cont_lat_row
                        GoSub PASTE_ROWS
                        cont_lat_row = cont_lat_row + 1
                    End If
                    
                    ActiveWorkbook.Save
                
                ElseIf (InStr(1, group, "p", vbTextCompare) > 0) Then
                
                    Windows("ptsd_lat.xls").Activate
                    
                    If (InStr(1, cond, "ea", vbTextCompare) > 0) Then
                        Sheets("ea").Select
                        PASTE_ROW = ptsd_lat_row
                        GoSub PASTE_ROWS
                    ElseIf (InStr(1, cond, "dt", vbTextCompare) > 0) Then
                        Sheets("dt").Select
                        PASTE_ROW = ptsd_lat_row
                        GoSub PASTE_ROWS
                    ElseIf (InStr(1, cond, "sa", vbTextCompare) > 0) Then
                        Sheets("sa").Select
                        PASTE_ROW = ptsd_lat_row
                        GoSub PASTE_ROWS
                    ElseIf (InStr(1, cond, "wm", vbTextCompare) > 0) Then
                        Sheets("wm").Select
                        PASTE_ROW = ptsd_lat_row
                        GoSub PASTE_ROWS
                        ptsd_lat_row = ptsd_lat_row + 1
                    End If
                    
                    ActiveWorkbook.Save
                
                End If
                
                Windows(foundfile).Activate
                ActiveWorkbook.Close
            
            Next i
        
            'Save and Close all peak files
            Windows("cont_elec.xls").Activate
            ActiveWorkbook.Close
            Windows("cont_amp.xls").Activate
            ActiveWorkbook.Close
            Windows("cont_lat.xls").Activate
            ActiveWorkbook.Close
            Windows("ptsd_elec.xls").Activate
            ActiveWorkbook.Close
            Windows("ptsd_amp.xls").Activate
            ActiveWorkbook.Close
            Windows("ptsd_lat.xls").Activate
            ActiveWorkbook.Close
            
        Else
            MsgBox "There were no files found."
        End If
    End With
    
    GoTo Fin
    
ADDSHEETS:
        Sheets(1).Name = ("ea")
        Sheets(2).Name = ("dt")
        Sheets(3).Name = ("sa")
        Sheets(3).Copy After:=Sheets(3)
        Sheets(4).Name = ("wm")
    Return
    
PASTE_ROWS:
        Cells(PASTE_ROW, 1).Select
        ActiveSheet.PASTE
        Cells(PASTE_ROW, 1).Insert Shift:=xlShiftToRight
        Cells(PASTE_ROW, 1).Value = foundfile
    Return
    
Fin:
    
End Sub

Sub ptsd_peaks_cleanup_all_files()
'
' ptsd_peaks_cleanup_all_files Macro
' Macro recorded 07/02/01 by Darren Weber
'

'   Gather all peak data for all peak files

    path = "D:\My Documents\THESIS\results\erp results\PEAKS"

    'Define and open all peak files
    cont_elec = path & "\cont_elec.xls"
    cont_amp = path & "\cont_amp.xls"
    cont_lat = path & "\cont_lat.xls"
    ptsd_elec = path & "\ptsd_elec.xls"
    ptsd_amp = path & "\ptsd_amp.xls"
    ptsd_lat = path & "\ptsd_lat.xls"
    
    Workbooks.Open cont_elec
    Workbooks.Open cont_amp
    Workbooks.Open cont_lat
    
    conditions = Array("ea", "dt", "sa", "wm")
    
    For cond = 0 To UBound(conditions)
    
        book = "c01_" & conditions(cond) & "_div.peaks.xls"
        Workbooks.Open (path & "\" & book)
        
        GoSub CopyHeading
        Windows("cont_elec.xls").Activate
        Sheets(cond + 1).Select
        GoSub InsertHeading
        
        GoSub CopyHeading
        Windows("cont_amp.xls").Activate
        Sheets(cond + 1).Select
        GoSub InsertHeading
        
        GoSub CopyHeading
        Windows("cont_lat.xls").Activate
        Sheets(cond + 1).Select
        GoSub InsertHeading
    
        Windows(book).Activate
        ActiveWorkbook.Close
    
    Next cond
    
    Windows("cont_elec.xls").Activate
    GoSub SaveClose
    Windows("cont_amp.xls").Activate
    GoSub SaveClose
    Windows("cont_lat.xls").Activate
    GoSub SaveClose

    Workbooks.Open ptsd_elec
    Workbooks.Open ptsd_amp
    Workbooks.Open ptsd_lat
    
    For cond = 0 To UBound(conditions)
    
        book = "p01_" & conditions(cond) & "_div.peaks.xls"
        Workbooks.Open (path & "\" & book)
        
        GoSub CopyHeading
        Windows("ptsd_elec.xls").Activate
        Sheets(cond + 1).Select
        GoSub InsertHeading
        
        GoSub CopyHeading
        Windows("ptsd_amp.xls").Activate
        Sheets(cond + 1).Select
        GoSub InsertHeading
        
        GoSub CopyHeading
        Windows("ptsd_lat.xls").Activate
        Sheets(cond + 1).Select
        GoSub InsertHeading
    
        Windows(book).Activate
        ActiveWorkbook.Close
        
    Next cond
    
    Windows("ptsd_elec.xls").Activate
    GoSub SaveClose
    Windows("ptsd_amp.xls").Activate
    GoSub SaveClose
    Windows("ptsd_lat.xls").Activate
    GoSub SaveClose
    
    GoTo Fin
    
CopyHeading:
        Windows(book).Activate
        Rows("1:2").Select
        Selection.Copy
    Return

InsertHeading:
        Rows("1:2").Select
        Selection.Insert Shift:=xlDown
        Range("A1:A2").Select
        Application.CutCopyMode = False
        Selection.Insert Shift:=xlToRight
        Application.CutCopyMode = xlCopy
        Cells.Select
        Cells.EntireColumn.AutoFit
        Cells(1, 1).Select
    Return
    
SaveClose:
        ActiveWorkbook.Save
        ActiveWorkbook.Close
    Return
    
Fin:

End Sub

Sub Twin_peaks_compose_all_files()
'
' Twin_peaks_compose_all_files Macro
' Macro recorded 25/10/00 by Darren Weber
'

'   Gather all peak data for all peak files

    path = "D:\My Documents\THESIS\results\erp results\PEAKS"

    'Define and open all peak files
    common_amp = "$path\common_sw_amp.xls"
    common_lat = "$path\common_sw_lat.xls"
    distractor_amp = "$path\distractor_sw_amp.xls"
    distractor_lat = "$path\distractor_sw_lat.xls"
    target_amp = "$path\target_sw_amp.xls"
    target_lat = "$path\target_sw_lat.xls"
    
    'Initialise row counters for placement of peak data
    common_amp_row = 279
    common_lat_row = 279
    distractor_amp_row = 279
    distractor_lat_row = 279
    target_amp_row = 279
    target_lat_row = 279
    'dist_com_amp_row = 279
    'dist_com_lat_row = 279
    'targ_com_amp_row = 279
    'targ_com_lat_row = 279
    
    Dim foundfile As String
    
    Set fs = Application.FileSearch
    With fs
        .LookIn = "D:\My Documents\THESIS\results\erp results\PEAKS"
        '.SearchSubFolders = True
        .FileName = "c*.peaks.xls"
        
        If .Execute > 0 Then
            
            Workbooks.Open common_amp
            Workbooks.Open common_lat
            Workbooks.Open distractor_amp
            Workbooks.Open distractor_lat
            Workbooks.Open target_amp
            Workbooks.Open target_lat
            
            'Workbooks.Open dist_com_amp
            'Workbooks.Open dist_com_lat
            'Workbooks.Open targ_com_amp
            'Workbooks.Open targ_com_lat
            
            'MsgBox "There are " & .FoundFiles.Count & " file(s) to process."
            
            For i = 1 To .FoundFiles.Count
                
                Workbooks.Open .FoundFiles(i)
                
                foundfile = right(.FoundFiles(i), 20)
                ' for dif waves use: foundfile = Right(.FoundFiles(i), 22)
                
                'Select Amplitude row (#3)
                Range(Cells(3, 1), Cells(3, 80)).Select
                Selection.Copy
                
                'Paste amp data into correct condition file
                If (InStr(1, foundfile, "_1.peaks.xls", vbTextCompare) > 0) Then
                    Windows("target_sw_amp.xls").Activate
                    Cells(target_amp_row, 1).Value = foundfile
                    Cells(target_amp_row, 2).Select
                    ActiveSheet.PASTE
                    ActiveWorkbook.Save
                    target_amp_row = target_amp_row + 1
                ElseIf (InStr(1, foundfile, "_2.peaks.xls", vbTextCompare) > 0) Then
                    Windows("distractor_sw_amp.xls").Activate
                    Cells(distractor_amp_row, 1).Value = foundfile
                    Cells(distractor_amp_row, 2).Select
                    ActiveSheet.PASTE
                    ActiveWorkbook.Save
                    distractor_amp_row = distractor_amp_row + 1
                ElseIf (InStr(1, foundfile, "_3.peaks.xls", vbTextCompare) > 0) Then
                    Windows("common_sw_amp.xls").Activate
                    Cells(common_amp_row, 1).Value = foundfile
                    Cells(common_amp_row, 2).Select
                    ActiveSheet.PASTE
                    ActiveWorkbook.Save
                    common_amp_row = common_amp_row + 1
                ElseIf (InStr(1, foundfile, "_1-3.peaks.xls", vbTextCompare) > 0) Then
                    Windows("targ-com_amp.xls").Activate
                    Cells(targ_com_amp_row, 1).Value = foundfile
                    Cells(targ_com_amp_row, 2).Select
                    ActiveSheet.PASTE
                    ActiveWorkbook.Save
                    targ_com_amp_row = targ_com_amp_row + 1
                ElseIf (InStr(1, foundfile, "_2-3.peaks.xls", vbTextCompare) > 0) Then
                    Windows("dist-com_amp.xls").Activate
                    Cells(dist_com_amp_row, 1).Value = foundfile
                    Cells(dist_com_amp_row, 2).Select
                    ActiveSheet.PASTE
                    ActiveWorkbook.Save
                    dist_com_amp_row = dist_com_amp_row + 1
                End If
                
                'Select latency row (#4)
                Windows(foundfile).Activate
                Range(Cells(4, 1), Cells(4, 80)).Select
                Selection.Copy
                
                'Paste lat data into correct condition file
                If (InStr(1, foundfile, "_1.peaks.xls", vbTextCompare) > 0) Then
                    Windows("target_sw_lat.xls").Activate
                    Cells(target_lat_row, 1).Value = foundfile
                    Cells(target_lat_row, 2).Select
                    ActiveSheet.PASTE
                    ActiveWorkbook.Save
                    target_lat_row = target_lat_row + 1
                ElseIf (InStr(1, foundfile, "_2.peaks.xls", vbTextCompare) > 0) Then
                    Windows("distractor_sw_lat.xls").Activate
                    Cells(distractor_lat_row, 1).Value = foundfile
                    Cells(distractor_lat_row, 2).Select
                    ActiveSheet.PASTE
                    ActiveWorkbook.Save
                    distractor_lat_row = distractor_lat_row + 1
                ElseIf (InStr(1, foundfile, "_3.peaks.xls", vbTextCompare) > 0) Then
                    Windows("common_sw_lat.xls").Activate
                    Cells(common_lat_row, 1).Value = foundfile
                    Cells(common_lat_row, 2).Select
                    ActiveSheet.PASTE
                    ActiveWorkbook.Save
                    common_lat_row = common_lat_row + 1
                ElseIf (InStr(1, foundfile, "_1-3.peaks.xls", vbTextCompare) > 0) Then
                    Windows("targ-com_lat.xls").Activate
                    Cells(targ_com_lat_row, 1).Value = foundfile
                    Cells(targ_com_lat_row, 2).Select
                    ActiveSheet.PASTE
                    ActiveWorkbook.Save
                    targ_com_lat_row = targ_com_lat_row + 1
                ElseIf (InStr(1, foundfile, "_2-3.peaks.xls", vbTextCompare) > 0) Then
                    Windows("dist-com_lat.xls").Activate
                    Cells(dist_com_lat_row, 1).Value = foundfile
                    Cells(dist_com_lat_row, 2).Select
                    ActiveSheet.PASTE
                    ActiveWorkbook.Save
                    dist_com_lat_row = dist_com_lat_row + 1
                End If
                
                Windows(foundfile).Activate
                ActiveWorkbook.Close
            
            Next i
        
            'Save and Close all peak files
            Windows("common_sw_amp.xls").Activate
            ActiveWorkbook.Close
            Windows("common_sw_lat.xls").Activate
            ActiveWorkbook.Close
            Windows("distractor_sw_amp.xls").Activate
            ActiveWorkbook.Close
            Windows("distractor_sw_lat.xls").Activate
            ActiveWorkbook.Close
            Windows("target_sw_amp.xls").Activate
            ActiveWorkbook.Close
            Windows("target_sw_lat.xls").Activate
            ActiveWorkbook.Close
            
            'Windows("dist-com_amp.xls").Activate
            'ActiveWorkbook.Close
            'Windows("dist-com_lat.xls").Activate
            'ActiveWorkbook.Close
            'Windows("targ-com_amp.xls").Activate
            'ActiveWorkbook.Close
            'Windows("targ-com_lat.xls").Activate
            'ActiveWorkbook.Close
            
        Else
            MsgBox "There were no files found."
        End If
    End With
    
End Sub

Sub Twin_Insert_N2_Columns()
'
' Twin_Insert_N2_Columns Macro
' Macro recorded 26/10/00 by Darren Weber
'
' Keyboard Shortcut: Ctrl+i
'
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight
    Columns("K:K").Select
    Selection.Insert Shift:=xlToRight
    ActiveWindow.SmallScroll ToRight:=8
    Columns("Q:Q").Select
    Selection.Insert Shift:=xlToRight
    Range("T1").Select
    ActiveWindow.SmallScroll ToRight:=9
    Columns("W:W").Select
    Selection.Insert Shift:=xlToRight
    Columns("AC:AC").Select
    Selection.Insert Shift:=xlToRight
    ActiveWindow.SmallScroll ToRight:=8
    Columns("AI:AI").Select
    Selection.Insert Shift:=xlToRight
    ActiveWindow.SmallScroll ToRight:=8
    Columns("AO:AO").Select
    Selection.Insert Shift:=xlToRight
    Range("AR1").Select
    ActiveWindow.SmallScroll ToRight:=7
    Columns("AU:AU").Select
    Selection.Insert Shift:=xlToRight
    Range("AX1").Select
    ActiveWindow.SmallScroll ToRight:=5
    Columns("BA:BA").Select
    Selection.Insert Shift:=xlToRight
    Range("BD1").Select
    ActiveWindow.SmallScroll ToRight:=4
    Columns("BG:BG").Select
    Selection.Insert Shift:=xlToRight
    Range("A1").Select
End Sub

Sub Twin_Cleanup_Subject_Numbers()
'
' Twin_Cleanup_Subject_Numbers Macro
' Macro recorded 1/11/00 by Darren Weber
'
' Keyboard Shortcut: Ctrl+n
'
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A2"), DataType:=xlFixedWidth, _
        OtherChar:=".", FieldInfo:=Array(Array(0, 1), Array(10, 1))
    Columns("A:A").EntireColumn.AutoFit
    Columns("B:B").Select
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
End Sub

