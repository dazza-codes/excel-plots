VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AxisProperties 
   Caption         =   "Axis Properties"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3855
   OleObjectBlob   =   "AxisProperties.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AxisProperties"
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()
    
    Dim ch As ChartObject
    For Each ch In ActiveSheet.ChartObjects
        
        If ch.Chart.HasAxis(xlValue) Then
            With ch.Chart.axes(xlValue)
                
                ERP_Plot70.AxisLine = .Border.LineStyle
                
                ERP_Plot70.majorticks = .MajorTickMark
                ERP_Plot70.minorticks = .MinorTickMark
                ERP_Plot70.labelpos = .TickLabelPosition
                
                With .TickLabels.font
                    ERP_Plot70.font = .Name
                    ERP_Plot70.font_style = .FontStyle
                    ERP_Plot70.font_size = .Size
                End With
            End With
            GoTo Init
        End If
    Next ch
    
Init:

    AxisProperties.ComboBoxAxisLine.ColumnCount = 1
    AxisProperties.ComboBoxAxisLine.ColumnWidths = 15
    AxisProperties.ComboBoxAxisLine.AddItem "None"
    AxisProperties.ComboBoxAxisLine.AddItem "Automatic"
    AxisProperties.ComboBoxAxisLine.AddItem "Continuous"
    AxisProperties.ComboBoxAxisLine.AddItem "Dash"
    AxisProperties.ComboBoxAxisLine.AddItem "DashDot"
    AxisProperties.ComboBoxAxisLine.AddItem "DashDotDot"
    AxisProperties.ComboBoxAxisLine.AddItem "Dot"
    AxisProperties.ComboBoxAxisLine.AddItem "Double"
    AxisProperties.ComboBoxAxisLine.AddItem "SlantDashDot"
    
    Select Case ERP_Plot70.AxisLine
        Case xlLineStyleNone
            AxisProperties.ComboBoxAxisLine.Value = "None"
        Case xlAutomatic
            AxisProperties.ComboBoxAxisLine.Value = "Automatic"
        Case xlContinuous
            AxisProperties.ComboBoxAxisLine.Value = "Continuous"
        Case xlDash
            AxisProperties.ComboBoxAxisLine.Value = "Dash"
        Case xlDashDot
            AxisProperties.ComboBoxAxisLine.Value = "DashDot"
        Case xlDashDotDot
            AxisProperties.ComboBoxAxisLine.Value = "DashDotDot"
        Case xlDot
            AxisProperties.ComboBoxAxisLine.Value = "Dot"
        Case xlDouble
            AxisProperties.ComboBoxAxisLine.Value = "Double"
        Case xlSlantDashDot
            AxisProperties.ComboBoxAxisLine.Value = "SlantDashDot"
    End Select
    
    AxisProperties.ComboBoxMajorTicks.ColumnCount = 1
    AxisProperties.ComboBoxMajorTicks.AddItem "None"
    AxisProperties.ComboBoxMajorTicks.AddItem "Cross"
    AxisProperties.ComboBoxMajorTicks.AddItem "Inside"
    AxisProperties.ComboBoxMajorTicks.AddItem "Outside"
    
    Select Case ERP_Plot70.majorticks
        Case xlTickMarkNone
            AxisProperties.ComboBoxMajorTicks.Value = "None"
        Case xlTickMarkCross
            AxisProperties.ComboBoxMajorTicks.Value = "Cross"
        Case xlTickMarkInside
            AxisProperties.ComboBoxMajorTicks.Value = "Inside"
        Case xlTickMarkOutside
            AxisProperties.ComboBoxMajorTicks.Value = "Outside"
    End Select

    AxisProperties.ComboBoxMinorTicks.ColumnCount = 1
    AxisProperties.ComboBoxMinorTicks.AddItem "None"
    AxisProperties.ComboBoxMinorTicks.AddItem "Cross"
    AxisProperties.ComboBoxMinorTicks.AddItem "Inside"
    AxisProperties.ComboBoxMinorTicks.AddItem "Outside"
    
    Select Case ERP_Plot70.minorticks
        Case xlTickMarkNone
            AxisProperties.ComboBoxMinorTicks.Value = "None"
        Case xlTickMarkCross
            AxisProperties.ComboBoxMinorTicks.Value = "Cross"
        Case xlTickMarkInside
            AxisProperties.ComboBoxMinorTicks.Value = "Inside"
        Case xlTickMarkOutside
            AxisProperties.ComboBoxMinorTicks.Value = "Outside"
    End Select
    
    AxisProperties.ComboBoxTickLabelPos.ColumnCount = 1
    AxisProperties.ComboBoxTickLabelPos.AddItem "None"
    AxisProperties.ComboBoxTickLabelPos.AddItem "Low"
    AxisProperties.ComboBoxTickLabelPos.AddItem "High"
    AxisProperties.ComboBoxTickLabelPos.AddItem "NextToAxis"
    
    Select Case ERP_Plot70.labelpos
        Case xlTickLabelPositionNone
            AxisProperties.ComboBoxTickLabelPos.Value = "None"
        Case xlTickLabelPositionLow
            AxisProperties.ComboBoxTickLabelPos.Value = "Low"
        Case xlTickLabelPositionHigh
            AxisProperties.ComboBoxTickLabelPos.Value = "High"
        Case xlTickLabelPositionNextToAxis
            AxisProperties.ComboBoxTickLabelPos.Value = "NextToAxis"
    End Select
    
    'AxisProperties.CommonDialogFonts.Flags = cdlCFPrinterFonts
    'AxisProperties.CommonDialogFonts.FontName = ERP_Plot70.font
    'AxisProperties.CommonDialogFonts.FontSize = ERP_Plot70.font_size
    'If InStr(1, ERP_Plot70.font_style, "Bold", vbTextCompare) > 0 Then
    '    AxisProperties.CommonDialogFonts.FontBold = True
    'End If
    'If InStr(1, ERP_Plot70.font_style, "Italic", vbTextCompare) > 0 Then
    '    AxisProperties.CommonDialogFonts.FontItalic = True
    'End If
    
End Sub

'Private Sub CommandButtonFonts_Click()
'    AxisProperties.CommonDialogFonts.ShowFont
'    ERP_Plot70.font = AxisProperties.CommonDialogFonts.FontName
'    ERP_Plot70.font_size = AxisProperties.CommonDialogFonts.FontSize
'    If (AxisProperties.CommonDialogFonts.FontBold And _
'        AxisProperties.CommonDialogFonts.FontItalic) Then
'        ERP_Plot70.font_style = "Bold Italic"
'        GoTo Fin
'    End If
'    If AxisProperties.CommonDialogFonts.FontBold Then
'        ERP_Plot70.font_style = "Bold"
'        GoTo Fin
'    End If
'    If AxisProperties.CommonDialogFonts.FontItalic Then
'        ERP_Plot70.font_style = "Italic"
'        GoTo Fin
'    End If
'    ERP_Plot70.font_style = "Regular"
'Fin:
'End Sub

Private Sub CommandOK_Click()
    
        Select Case AxisProperties.ComboBoxAxisLine.Value
        Case "None"
            ERP_Plot70.AxisLine = xlLineStyleNone
        Case "Automatic"
            ERP_Plot70.AxisLine = xlAutomatic
        Case "Continuous"
            ERP_Plot70.AxisLine = xlContinuous
        Case "Dash"
            ERP_Plot70.AxisLine = xlDash
        Case "DashDot"
            ERP_Plot70.AxisLine = xlDashDot
        Case "DashDotDot"
            ERP_Plot70.AxisLine = xlDashDotDot
        Case "Dot"
            ERP_Plot70.AxisLine = xlDot
        Case "Double"
            ERP_Plot70.AxisLine = xlDouble
        Case "SlantDashDot"
            ERP_Plot70.AxisLine = xlSlantDashDot
    End Select
    
    Select Case AxisProperties.ComboBoxMajorTicks.Value
        Case "None"
            ERP_Plot70.majorticks = xlTickMarkNone
        Case "Cross"
            ERP_Plot70.majorticks = xlTickMarkCross
        Case "Inside"
            ERP_Plot70.majorticks = xlTickMarkInside
        Case "Outside"
            ERP_Plot70.majorticks = xlTickMarkOutside
    End Select
    
    Select Case AxisProperties.ComboBoxMinorTicks.Value
        Case "None"
            ERP_Plot70.minorticks = xlTickMarkNone
        Case "Cross"
            ERP_Plot70.minorticks = xlTickMarkCross
        Case "Inside"
            ERP_Plot70.minorticks = xlTickMarkInside
        Case "Outside"
            ERP_Plot70.minorticks = xlTickMarkOutside
    End Select
    
    Select Case AxisProperties.ComboBoxTickLabelPos.Value
        Case "None"
            ERP_Plot70.labelpos = xlTickLabelPositionNone
        Case "Low"
            ERP_Plot70.labelpos = xlTickLabelPositionLow
        Case "High"
            ERP_Plot70.labelpos = xlTickLabelPositionHigh
        Case "NextToAxis"
            ERP_Plot70.labelpos = xlTickLabelPositionNextToAxis
    End Select
    
    ERP_Plot70.bRun = True
    Unload Me
End Sub

Private Sub CommandCancel_Click()
    ERP_Plot70.bRun = False
    Unload Me
End Sub
