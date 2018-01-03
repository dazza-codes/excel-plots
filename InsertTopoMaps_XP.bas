Attribute VB_Name = "InsertTopoMaps"
Sub InsertScalpTopography()
Attribute InsertScalpTopography.VB_Description = "Macro recorded 12/09/2002 by Numerous"
Attribute InsertScalpTopography.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.InsertScalpTopography"
'
' InsertScalpTopography Macro
' Macro recorded 12/09/2002 by Numerous
'
    
    ' switch between "topomaps" or "topocontours" for associated figures
    
    Dim FName As String
    
    Studies = Array("ea", "sa", "wm")
    
    DATTypes = Array("link14hz", "scd14hz")
    
    For s = 0 To UBound(Studies)
        
        Study = Studies(s)
        
        Select Case Study
        Case "ea"
            Components = Array("coat", "coac", "coat-oac", "poat", "poac", "poat-oac")
        Case "sa"
            Components = Array("coac", "couc", "coac-ouc", "poac", "pouc", "poac-ouc")
        Case "wm"
            Components = Array("ctac", "coac", "ctac-oac", "ptac", "poac", "ptac-oac")
        End Select
        
        For c = 0 To UBound(Components)
            
            Comp = Components(c)
            
            For d = 0 To UBound(DATTypes)
                
                DATType = DATTypes(d)
                
                DOCPath = "E:\data_emse\ptsdpet\grand_mean\"
                IMGPath = "E:\data_emse\ptsdpet\grand_mean\" & DATType & "\topocontours\" & Study & "\tmp\"
                IMGName = Comp & "_" & DATType
                IMGType = ".png"
                
                DOCName = UCase(Study) & "_topo_" & DATType & "_" & Comp & ".doc"
                
                ' --- Testing ---
                't = 50
                'LAT = Format(t, "0#####.00")
                'FName = IMGPath & IMGName & Views(0) & LAT & IMGType
                ' Example FName:
                ' "E:\data_emse\ptsdpet\grand_mean\link14hz\topomaps_ea\tmp\coac_link14hz_front_00050.00.png"
                'ans = MsgBox(FName, vbOKCancel)
                'End
                
                
                ' Save a new document to a new file name "DOCName"
                Documents.Add Template:="Normal", NewTemplate:=False, DocumentType:=0
                
                ActiveDocument.SaveAs FileName:=DOCPath & DOCName, FileFormat:= _
                wdFormatDocument, LockComments:=False, Password:="", AddToRecentFiles:= _
                True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:= _
                False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
                SaveAsAOCELetter:=False
                
                ' Clear the current document of all content
                Selection.WholeStory
                Selection.Delete Unit:=wdCharacter, Count:=1
                ActiveDocument.Save
                
                ' Ensure page setup is A4 with 1cm margins
                With ActiveDocument.PageSetup
                    .LineNumbering.Active = False
                    .Orientation = wdOrientPortrait
                    .TopMargin = CentimetersToPoints(1)
                    .BottomMargin = CentimetersToPoints(1)
                    .LeftMargin = CentimetersToPoints(1)
                    .RightMargin = CentimetersToPoints(1)
                    .Gutter = CentimetersToPoints(0)
                    .HeaderDistance = CentimetersToPoints(1.25)
                    .FooterDistance = CentimetersToPoints(1.25)
                    .PageWidth = CentimetersToPoints(21)
                    .PageHeight = CentimetersToPoints(29.7)
                    .FirstPageTray = wdPrinterDefaultBin
                    .OtherPagesTray = wdPrinterDefaultBin
                    .SectionStart = wdSectionNewPage
                    .OddAndEvenPagesHeaderFooter = False
                    .DifferentFirstPageHeaderFooter = False
                    .VerticalAlignment = wdAlignVerticalTop
                    .SuppressEndnotes = False
                    .MirrorMargins = False
                    .TwoPagesOnOne = False
                    .BookFoldPrinting = False
                    .BookFoldRevPrinting = False
                    .BookFoldPrintingSheets = 1
                    .GutterPos = wdGutterPosLeft
                End With
                
                ' Define the views
                'Views = Array("_front_", "_back_", "_left_", "_right_")
                
                For t = 50 To 800 Step 5
                    
                    LAT = Format(t, "0####.00")
                    
                    FName = IMGName & "*" & LAT & IMGType
                    'MsgBox FName
                    
                    Set fs = Application.FileSearch
                    With fs
                        .LookIn = IMGPath
                        .FileName = FName
                        If .Execute > 0 Then
                            'MsgBox "There were " & .FoundFiles.Count & " file(s) found."
                            
                            ' Front View
                            FName = .FoundFiles(2)
                            Selection.TypeText Text:=t & vbTab
                            Selection.InlineShapes.AddPicture FileName:=FName _
                                , LinkToFile:=False, SaveWithDocument:=True
                            
                            ' Back View
                            FName = .FoundFiles(1)
                            Selection.TypeText Text:=vbTab
                            Selection.InlineShapes.AddPicture FileName:=FName _
                                , LinkToFile:=False, SaveWithDocument:=True
                            
                            ' Left View
                            FName = .FoundFiles(3)
                            Selection.TypeText Text:=vbTab
                            Selection.InlineShapes.AddPicture FileName:=FName _
                                , LinkToFile:=False, SaveWithDocument:=True
                            
                            ' Right View
                            FName = .FoundFiles(4)
                            Selection.TypeText Text:=vbTab
                            Selection.InlineShapes.AddPicture FileName:=FName _
                                , LinkToFile:=False, SaveWithDocument:=True
                            Selection.TypeParagraph
                            
                        Else
                            MsgBox "There were no files found."
                            End
                        End If
                    End With
                    
                Next t
                
                ActiveDocument.Save
                ActiveDocument.Close
                
            Next ' Components
        Next ' DATTypes
    Next ' Studies
    
End Sub

Sub ComponentScalpTopography()
'
' InsertScalpTopography Macro
' Macro recorded 12/09/2002 by Numerous
'
    
    ' switch between "topomaps" or "topocontours" for associated figures
    
    
    
    Studies = Array("ea", "sa", "wm")
    'Studies = Array("wm")
    
    DIFF = Array("", "_dif")
    'DIFF = Array("_dif")
    
    DATTypes = Array("link14hz", "scd14hz")
    'DATTypes = Array("link14hz")
    
    For s = 0 To UBound(Studies)
        
        Study = Studies(s)
        
        For Cdif = 0 To UBound(DIFF)
            
            DIF = DIFF(Cdif)
            
            Select Case Study
            Case "ea"
                If StrComp(DIF, "_dif", vbTextCompare) = 0 Then
                    Components = Array("coat-oac", "poat-oac")
                    times_link = Array(320, 340, 495, 650, 700)
                    times_scd = Array(330, 480, 575)
                Else
                    Components = Array("coat", "coac", "poat", "poac")
                    times_link = Array(85, 155, 240, 300, 350, 450, 520, 700)
                    times_scd = Array(145, 175, 240, 275, 360, 460, 575)
                End If
                
            Case "sa"
                If StrComp(DIF, "_dif", vbTextCompare) = 0 Then
                    Components = Array("coac-ouc", "poac-ouc")
                    times_link = Array(280, 450)
                    times_scd = Array(70, 135, 235, 260, 295, 400, 630, 710)
                Else
                    Components = Array("coac", "couc", "poac", "pouc")
                    times_link = Array(80, 100, 150, 200, 250, 400)
                    times_scd = Array(60, 85, 120, 150, 180, 230, 250, 350, 380, 420, 450, 525)
                End If
                
            Case "wm"
                If StrComp(DIF, "_dif", vbTextCompare) = 0 Then
                    Components = Array("ctac-oac", "ptac-oac")
                    times_link = Array(200, 330, 575)
                    times_scd = Array(370, 580)
                Else
                    Components = Array("ctac", "coac", "ptac", "poac")
                    times_link = Array(85, 95, 155, 250, 300, 410, 530)
                    times_scd = Array(90, 115, 140, 250, 350, 500)
                End If
                
            End Select
            
            For d = 0 To UBound(DATTypes)
                
                DATType = DATTypes(d)
                
                
                Select Case DATType
                Case "link14hz"
                    times = times_link
                Case "scd14hz"
                    times = times_scd
                End Select
                
                
                DOCName = UCase(Study) & "_topo_" & DATType & DIF & "_components.doc"
                
                DOCPath = "E:\data_emse\ptsdpet\grand_mean\"
                
                IMGPath = "E:\data_emse\ptsdpet\grand_mean\" & DATType & "\topocontours\" & Study & "\tmp\"
                ColorBarPath = "E:\data_emse\ptsdpet\grand_mean\" & DATType & "\topocontours\"
                
                IMGType = ".png"
                
                ' Save the current document to a new file name "DOCName"
                
                Documents.Add Template:="Normal", NewTemplate:=False, DocumentType:=0
                
                ActiveDocument.SaveAs FileName:=DOCPath & DOCName, FileFormat:= _
                wdFormatDocument, LockComments:=False, Password:="", AddToRecentFiles:= _
                True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:= _
                False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
                SaveAsAOCELetter:=False
                
                ' Clear the current document of all content
                Selection.WholeStory
                Selection.Delete Unit:=wdCharacter, Count:=1
                ActiveDocument.Save
                
                ' Ensure page setup is A4 with 1cm margins
                With ActiveDocument.PageSetup
                    .LineNumbering.Active = False
                    .Orientation = wdOrientPortrait
                    .TopMargin = CentimetersToPoints(1)
                    .BottomMargin = CentimetersToPoints(1)
                    .LeftMargin = CentimetersToPoints(1)
                    .RightMargin = CentimetersToPoints(1)
                    .Gutter = CentimetersToPoints(0)
                    .HeaderDistance = CentimetersToPoints(1.25)
                    .FooterDistance = CentimetersToPoints(1.25)
                    .PageWidth = CentimetersToPoints(21)
                    .PageHeight = CentimetersToPoints(29.7)
                    .FirstPageTray = wdPrinterDefaultBin
                    .OtherPagesTray = wdPrinterDefaultBin
                    .SectionStart = wdSectionNewPage
                    .OddAndEvenPagesHeaderFooter = False
                    .DifferentFirstPageHeaderFooter = False
                    .VerticalAlignment = wdAlignVerticalTop
                    .SuppressEndnotes = False
                    .MirrorMargins = False
                    .TwoPagesOnOne = False
                    .BookFoldPrinting = False
                    .BookFoldRevPrinting = False
                    .BookFoldPrintingSheets = 1
                    .GutterPos = wdGutterPosLeft
                End With
                
                
                For t = 0 To UBound(times)
                    
                    'ans = MsgBox(Times(n), vbOKCancel)
                    
                    LAT = Format(times(t), "0####.00")
                    
                    For c = 0 To UBound(Components)
                        
                        IMGName = Components(c) & "_" & DATType
                        
                        FName = IMGName & "*" & LAT & IMGType
                        'MsgBox FName
                        
                        Set fs = Application.FileSearch
                        With fs
                            .LookIn = IMGPath
                            .FileName = FName
                            If .Execute > 0 Then
                                'MsgBox "There were " & .FoundFiles.Count & " file(s) found."
                                
                                ' Front View
                                FName = .FoundFiles(2)
                                Selection.TypeText Text:=Components(c) & "_" & times(t) & vbTab
                                Selection.InlineShapes.AddPicture FileName:=FName _
                                    , LinkToFile:=False, SaveWithDocument:=True
                                
                                ' Back View
                                FName = .FoundFiles(1)
                                Selection.TypeText Text:=vbTab
                                Selection.InlineShapes.AddPicture FileName:=FName _
                                    , LinkToFile:=False, SaveWithDocument:=True
                                
                                ' Left View
                                FName = .FoundFiles(3)
                                Selection.TypeText Text:=vbTab
                                Selection.InlineShapes.AddPicture FileName:=FName _
                                    , LinkToFile:=False, SaveWithDocument:=True
                                
                                ' Right View
                                FName = .FoundFiles(4)
                                Selection.TypeText Text:=vbTab
                                Selection.InlineShapes.AddPicture FileName:=FName _
                                    , LinkToFile:=False, SaveWithDocument:=True
                                Selection.TypeParagraph
                                
                            Else
                                MsgBox "There were no files found."
                                End
                            End If
                        End With
                    Next c 'component
                Next t 'time
                
                Rows = UBound(times) * UBound(Components)
                
                Selection.WholeStory
                Selection.ConvertToTable Separator:=wdSeparateByTabs, NumColumns:=5, _
                    NumRows:=Rows, AutoFitBehavior:=wdAutoFitContent
                With Selection.Tables(1)
                    .Style = "Table Grid"
                    .ApplyStyleHeadingRows = True
                    .ApplyStyleLastRow = True
                    .ApplyStyleFirstColumn = True
                    .ApplyStyleLastColumn = True
                End With
                
                Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
                Selection.HomeKey Unit:=wdStory
                
                Selection.InsertRowsAbove 1
                Selection.Rows.ConvertToText Separator:=wdSeparateByTabs, NestedTables:=True
                Selection.HomeKey Unit:=wdLine
                
                CBarName = ColorBarPath & UCase(Study) & "_" & DATType & DIF & "_colorbar.png"
                
                Selection.InlineShapes.AddPicture FileName:=CBarName, LinkToFile:=False, _
                                                  SaveWithDocument:=True
                
                Selection.HomeKey Unit:=wdStory
                
                ActiveDocument.Save
                ActiveDocument.Close
                
            Next d 'data type
        Next Cdif
    Next s 'study
    
End Sub

Sub InsertCorticalSourceMaps()
'
' InsertCorticalSourceMaps Macro
' Macro recorded 12/09/2002 by Numerous
'
    
    
    Dim FName As String
    
    Studies = Array("ea", "sa", "wm")
    
    Groups = Array("cont", "ptsd", "tmap")
    
    For s = 0 To UBound(Studies)
        
        Study = Studies(s)
        
        Select Case Study
        Case "ea"
            Components = Array("oat", "oac", "dif")
            times = Array(80, 145, 240, 300, 350, 450, 550)
        Case "sa"
            Components = Array("oac", "ouc", "dif")
            times = Array(100, 150, 180, 250, 400)
        Case "wm"
            Components = Array("tac", "oac", "dif")
            times = Array(80, 90, 150, 260, 300, 400, 550)
        End Select
        
        For c = 0 To UBound(Components)
            
            Comp = Components(c)
            
            DOCPath = "D:\freesurfer\subjects\"
            IMGPath = "D:\freesurfer\subjects\glm_means\tmp\"
            
            IMGName = "rh.cortex." & Study
            IMGType = ".png"
            
            DOCName = UCase(Study) & "_cortex_" & Comp & ".doc"
            
            ' Save a new document to a new file name "DOCName"
            Documents.Add Template:="Normal", NewTemplate:=False, DocumentType:=0
            
            ActiveDocument.SaveAs FileName:=DOCPath & DOCName, FileFormat:= _
            wdFormatDocument, LockComments:=False, Password:="", AddToRecentFiles:= _
            True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:= _
            False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
            SaveAsAOCELetter:=False
            
            ' Clear the current document of all content
            Selection.WholeStory
            Selection.Delete Unit:=wdCharacter, Count:=1
            ActiveDocument.Save
            
            For t = 0 To UBound(times)
                
                'ans = MsgBox(times(t), vbOKCancel)
                
                LAT = Format(times(t), "0###")
                
                'rh.cortex.ea_0080_dif_cont.ant.png
                
                FName = IMGName & "_" & LAT & "_" & Comp & "_" & Groups(0) & "*" & IMGType
                'MsgBox FName
                
                Set fs = Application.FileSearch
                With fs
                    .LookIn = IMGPath
                    .FileName = FName
                    If .Execute > 0 Then
                        
                        'MsgBox "There were " & .FoundFiles.Count & " file(s) found."
                        'For f = 1 To .FoundFiles.Count
                        '    MsgBox "FoundFile " & f & " = " & .FoundFiles(f)
                        'Next f
                        
                        CName = Replace(FName, "rh.cortex.", "")
                        CName = Replace(CName, "_cont*.png", "")
                        
                        Selection.TypeText Text:=CName
                        Selection.TypeParagraph
                        Selection.TypeText Text:=vbTab
                        
                        For g = 0 To UBound(Groups)
                            ' Table headers
                            Selection.TypeText Text:=Groups(g)
                            Selection.TypeText Text:=vbTab
                        Next g
                        Selection.TypeParagraph
                        Selection.TypeText Text:=vbTab
                        
                        For g = 0 To UBound(Groups)
                            ' Front View
                            FName = Replace(.FoundFiles(1), Groups(0), Groups(g))
                            Selection.InlineShapes.AddPicture FileName:=FName _
                                , LinkToFile:=False, SaveWithDocument:=True
                            Selection.TypeText Text:=vbTab
                        Next g
                        Selection.TypeParagraph
                        Selection.TypeText Text:=vbTab
                        
                        For g = 0 To UBound(Groups)
                            ' Back View
                            FName = Replace(.FoundFiles(5), Groups(0), Groups(g))
                            Selection.InlineShapes.AddPicture FileName:=FName _
                                , LinkToFile:=False, SaveWithDocument:=True
                            Selection.TypeText Text:=vbTab
                        Next g
                        Selection.TypeParagraph
                        Selection.TypeText Text:=vbTab
                        
                        For g = 0 To UBound(Groups)
                            ' Left View
                            FName = Replace(.FoundFiles(4), Groups(0), Groups(g))
                            Selection.InlineShapes.AddPicture FileName:=FName _
                                , LinkToFile:=False, SaveWithDocument:=True
                            Selection.TypeText Text:=vbTab
                        Next g
                        Selection.TypeParagraph
                        Selection.TypeText Text:=vbTab
                        
                        For g = 0 To UBound(Groups)
                            ' Right View
                            FName = Replace(.FoundFiles(6), Groups(0), Groups(g))
                            Selection.InlineShapes.AddPicture FileName:=FName _
                                , LinkToFile:=False, SaveWithDocument:=True
                            Selection.TypeText Text:=vbTab
                        Next g
                        Selection.TypeParagraph
                        Selection.TypeText Text:=vbTab
                        
                        For g = 0 To UBound(Groups)
                            ' Inferior View
                            FName = Replace(.FoundFiles(3), Groups(0), Groups(g))
                            Selection.InlineShapes.AddPicture FileName:=FName _
                                , LinkToFile:=False, SaveWithDocument:=True
                            Selection.TypeText Text:=vbTab
                        Next g
                        Selection.TypeParagraph
                        Selection.TypeText Text:=vbTab
                        
                        For g = 0 To UBound(Groups)
                            ' Superior View
                            FName = Replace(.FoundFiles(7), Groups(0), Groups(g))
                            Selection.InlineShapes.AddPicture FileName:=FName _
                                , LinkToFile:=False, SaveWithDocument:=True
                            Selection.TypeText Text:=vbTab
                        Next g
                        Selection.TypeParagraph
                        
                        ' Convert to table
                        Selection.MoveUp Unit:=wdLine, Count:=9, Extend:=wdExtend
                        Selection.ConvertToTable Separator:=wdSeparateByTabs, NumColumns:=5, _
                            NumRows:=7, AutoFitBehavior:=wdAutoFitContent
                        
                        ' Add Means Colorbar to first column
                        Selection.MoveLeft Unit:=wdCharacter, Count:=1
                        Selection.MoveDown Unit:=wdLine, Count:=1
                        Selection.MoveDown Unit:=wdLine, Count:=5, Extend:=wdExtend
                        Selection.Cells.Merge
                        FName = .FoundFiles(2)
                        Selection.InlineShapes.AddPicture FileName:=FName _
                            , LinkToFile:=False, SaveWithDocument:=True
                        
                        ' Add T-test Colorbar to last column
                        Selection.MoveUp Unit:=wdLine, Count:=1
                        Selection.MoveRight Unit:=wdWord, Count:=7
                        Selection.MoveDown Unit:=wdLine, Count:=1
                        Selection.MoveDown Unit:=wdLine, Count:=5, Extend:=wdExtend
                        Selection.Cells.Merge
                        FName = Replace(.FoundFiles(2), Groups(0), "tmap")
                        Selection.InlineShapes.AddPicture FileName:=FName _
                            , LinkToFile:=False, SaveWithDocument:=True
                        
                        
                        Selection.MoveDown Unit:=wdLine, Count:=1
                        Selection.TypeParagraph
                        Selection.InsertBreak Type:=wdPageBreak
                        
                    Else
                        MsgBox "There were no files found."
                        End
                    End If
                End With
            Next t
            
            ' Call subfunction...
            FormatTables
            
            ' Ensure page setup is A4 with 1cm margins
            With ActiveDocument.PageSetup
                .LineNumbering.Active = False
                .Orientation = wdOrientPortrait
                .TopMargin = CentimetersToPoints(1)
                .BottomMargin = CentimetersToPoints(1)
                .LeftMargin = CentimetersToPoints(1)
                .RightMargin = CentimetersToPoints(1)
                .Gutter = CentimetersToPoints(0)
                .HeaderDistance = CentimetersToPoints(1.25)
                .FooterDistance = CentimetersToPoints(1.25)
                .PageWidth = CentimetersToPoints(21)
                .PageHeight = CentimetersToPoints(29.7)
                .FirstPageTray = wdPrinterDefaultBin
                .OtherPagesTray = wdPrinterDefaultBin
                .SectionStart = wdSectionNewPage
                .OddAndEvenPagesHeaderFooter = False
                .DifferentFirstPageHeaderFooter = False
                .VerticalAlignment = wdAlignVerticalTop
                .SuppressEndnotes = False
                .MirrorMargins = False
                .TwoPagesOnOne = False
                .BookFoldPrinting = False
                .BookFoldRevPrinting = False
                .BookFoldPrintingSheets = 1
                .GutterPos = wdGutterPosLeft
            End With
            
            ActiveDocument.Save
            ActiveDocument.Close
        Next c
    Next ' Studies
    
End Sub
Sub FormatTables()
Attribute FormatTables.VB_Description = "Macro recorded 2/14/2003 by Darren L. Weber"
Attribute FormatTables.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro3"
'
' FormatTable Macro
' Macro recorded 2/14/2003 by Darren L. Weber
'

    For Each Table In ActiveDocument.Tables
        
        Table.Select
        With Table
            .TopPadding = CentimetersToPoints(0)
            .BottomPadding = CentimetersToPoints(0)
            .LeftPadding = CentimetersToPoints(0.1)
            .RightPadding = CentimetersToPoints(0.1)
            .Spacing = 0
            .AllowPageBreaks = True
            .AllowAutoFit = True
        End With
        With Selection.Cells(1)
            .WordWrap = True
            .FitText = False
        End With
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Next Table
    
End Sub
