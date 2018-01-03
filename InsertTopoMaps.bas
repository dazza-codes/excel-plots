Attribute VB_Name = "ScalpTopography"
Sub InsertScalpTopography()
Attribute InsertScalpTopography.VB_Description = "Macro recorded 12/09/2002 by Numerous"
Attribute InsertScalpTopography.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.InsertScalpTopography"
'
' InsertScalpTopography Macro
' Macro recorded 12/09/2002 by Numerous
'
    
    
    Dim FName As String
    
    DATType = "link"
    
    Study = "sa"
    Comp = "poac-ouc"
    
    IMGPath = "E:\data_emse\ptsdpet\grand_mean\" & DATType & "14hz\topomaps\" & Study & "\tmp\"
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
    
    ActiveDocument.SaveAs FileName:=DOCName, FileFormat:= _
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
    
End Sub

Sub ComponentScalpTopography()
'
' ComponentScalpTopography Macro
' Macro recorded 12/09/2002 by Numerous
'
    
    
    ' LINK SA Times
    'times = [  60  80 120; % first negative peak N80 (SC?)
    '           60  90 120; % first positive peak P100 (OT)
    '          100 150 200; % positive & negative peak P150 (SF) / N150 (OT)
    '          150 200 250; % negative N200 @ PT,AT and SP,SC
    '          190 250 300; % positive P250 (SP)
    '          300 400 500 ]; % P400 (SC,SP)
    'times = Array(100, 150, 200, 250, 400)
    'Comp = Array("coac", "couc", "poac", "pouc")
    'DIF = ""
    'DATType = "link"
    'Study = "sa"
    
    ' LINK SA DIF
    'times = [  80 100 150;
    '           80 125 160;
    '          140 190 250;
    '          180 260 350;  % ND250 (OT, PT, SP)
    '          350 440 550]; % PD450 (SPF, SF, SC)
    'times = Array(280, 450) ' 280/450 is mean latency
    'Comp = Array("coac-ouc", "poac-ouc")
    'DIF = "_dif"
    'DATType = "link"
    'Study = "sa"
    
    ' SCD SA
    '  40  85 160; % N85 @ IC/SC, see N120 below
    ' 120 180 240; % N180 @ PF & PT
    ' 125 230 300; % N230 @ PT -> IP/SP & SC/SF
    ' 250 380 420; % N380 @ PT, IPF
    ' 400 525 600; % N525 @ IPF
    '  40  60 100; % P60 @ OT
    '  80 120 180; % P120 @ PT & PF, N120/N140 @ OT
    '  80 150 180; % P150 @ IP/SP -> Lateral IF
    ' 180 250 350; % P250 @ OT & IPF/SPF
    ' 300 350 380; % P350 @ SP/IP
    ' 375 420 525; % P420 @ IF/SF & SC
    ' 300 450 525; % P450 @ SP
    'times = Array(60, 85, 120, 150, 180, 230, 250, 350, 380, 420, 450, 525)
    'Comp = Array("coac", "couc", "poac", "pouc")
    'DIF = ""
    'DATType = "scd"
    'Study = "sa"
    
    ' SCD SA DIF
    '       40  70 100; % PD70 @ OT/IP & IC/IF (L>R)
    '      120 135 150; % ND135 @ PT & IPF
    '      200 235 260; % PD235 @ IPF & OT
    '      200 295 340; % PD295 @ OT/SP, IF/AT
    '      200 260 400; % ND260 @ SC/SF, & PT @ 325 ms
    '      350 400 550; % PD400 @ SP/IF,SF/AT (L>R)
    '      510 630 750; % ND630 @ IPF/SPF & PT (R < L)
    '      625 710 750; % PD710 @ IC, AT (L > R)
    'times = Array(70, 135, 235, 260, 295, 400, 630, 710)
    'Comp = Array("coac-ouc", "poac-ouc")
    'DIF = "_dif"
    'DATType = "scd"
    'Study = "sa"
    
    
    
    ' LINK EA
    'times = [  50  80 110;  % N80  @ SF        , P80  @ OT/PT
    '      100 145 200;  % N150 @ OT/PT     , P150 @ SPF (SC)
    '      210 240 270;  %                  , P240 @ SPF/IPF & OT/SP
    '      250 300 350;  % N300 @ L_PT/OT/IP,
    '      250 350 625;  % Series: P350 @ Frontal (IPF), P500 @ SP
    '      250 450 625;  % Series: P350 @ Frontal (IPF), P500 @ SP
    '      250 550 625;  % Series: P350 @ Frontal (IPF), P500 @ SP; Also N550 @ left frontal
    '      600 700 800;  % P700 @ R_SC/SP
    '    ];
    ' LINK EA DIF
    'times = [ 220 300 420;  % ND300, Left PT/IP,SP/IC,SC @ 300 - 350
    '          250 350 400;  % PD350, Prefrontal
    '          420 510 580;  % PD500, Parietal
    '          450 650 800;  % ND650, Frontal
    '          620 700 800;  % PD700, Parietal
    '        ];
    ' LINK WM
    'times = [  60  80 100;
    '           70  90 110;
    '          100 150 200;
    '          200 260 300;
    '          250 300 350;
    '          300 400 480;
    '          450 550 650 ];
    ' LINK WM DIF
    'times = [ 100 200 250; %ND200
    '          250 300 450; %ND300
    '          300 550 700 ];
    times = Array(170, 200, 330, 300, 550)
    Comp = Array("ctac-oac", "ptac-oac")
    DIF = "_dif"
    DATType = "link"
    Study = "wm"
    
    
    
    DOCName = UCase(Study) & "_topo_" & DATType & DIF & "_components.doc"
    
    IMGPath = "\\POTZII\data2\data_emse\ptsdpet\grand_mean\" & DATType & "14hz\topomaps\" & Study & "\tmp\"
    ColorBarPath = "\\POTZII\data2\data_emse\ptsdpet\grand_mean\" & DATType & "14hz\topomaps\"
    
    
    'IMGPath = "E:\data_emse\ptsdpet\grand_mean\" & DATType & "14hz\topomaps\" & Study & "\tmp\"
    'ColorBarPath = "E:\data_emse\ptsdpet\grand_mean\" & DATType & "14hz\topomaps\"
    
    IMGType = ".png"
    
    ' Create a new document
    Documents.Add 'Template:="Normal", NewTemplate:=False, DocumentType:=0
    
    ' Save the current document to a new file name "DOCName"
    ActiveDocument.SaveAs FileName:=DOCName, FileFormat:= _
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
        '.BookFoldPrinting = False
        '.BookFoldRevPrinting = False
        '.BookFoldPrintingSheets = 1
        '.GutterPos = wdGutterPosLeft
    End With
    
    
    For t = 0 To UBound(times)
        
        'ans = MsgBox(Times(n), vbOKCancel)
        
        LAT = Format(times(t), "0####.00")
        
        For c = 0 To UBound(Comp)
            
            IMGName = Comp(c) & "_" & DATType
            
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
                    Selection.TypeText Text:=Comp(c) & "_" & times(t) & vbTab
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
        Next c
    Next t
    
    Rows = UBound(times) * UBound(Comp)
    
    Selection.WholeStory
    Selection.ConvertToTable Separator:=wdSeparateByTabs, NumColumns:=5, _
        NumRows:=Rows ', AutoFitBehavior:=wdAutoFitContent
    With Selection.Tables(1)
        '.Style = "Table Grid"
        '.ApplyStyleHeadingRows = True
        '.ApplyStyleLastRow = True
        '.ApplyStyleFirstColumn = True
        '.ApplyStyleLastColumn = True
    End With
    'Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
    Selection.HomeKey Unit:=wdStory
    
    For t = 0 To UBound(times)
        Selection.MoveDown Unit:=wdLine, Count:=UBound(Comp) + 1
        Selection.InsertRows (1)
        Selection.Rows.ConvertToText Separator:=wdSeparateByTabs ', NestedTables:=True
        Selection.HomeKey Unit:=wdLine
        Selection.InsertBreak Type:=wdPageBreak
        Selection.EndKey Unit:=wdLine, Extend:=wdExtend
        Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
        Selection.Delete Unit:=wdCharacter, Count:=1
        
        Selection.MoveUp Unit:=wdLine, Count:=6
        Selection.EndKey Unit:=wdLine
        Selection.MoveRight Unit:=wdCharacter, Count:=8
        Selection.InsertColumns 'Right
        Selection.Cells.Merge
        
        CBarName = ColorBarPath & UCase(Study) & "_" & DATType & DIF & "_colorbar.png"
        
        Selection.InlineShapes.AddPicture FileName:=CBarName, LinkToFile:=False, SaveWithDocument:=True
        Selection.MoveRight Unit:=wdCharacter, Count:=2
        Selection.MoveDown Unit:=wdLine, Count:=6
        
        Selection.MoveRight Unit:=wdCharacter, Count:=1
    Next t
    
    ActiveDocument.Save
    
End Sub

