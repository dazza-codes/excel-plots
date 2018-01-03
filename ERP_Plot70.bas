Attribute VB_Name = "ERP_Plot70"
'Declare public variables, set by Forms
Public dMin As Double, _
       dMax As Double, _
       dMajor As Double, _
       dMinor As Double, _
       dCrossesAt As Double, _
       dScaleXk As Double, _
       dScaleYk As Double, _
       lTickMarkSp As Long, _
       lTickLabelSp As Long, _
       lScaleXOrigin As Long, _
       lScaleYOrigin As Long, _
       bXreverse As Boolean, _
       bYreverse As Boolean, _
       bYBetween As Boolean, _
       bRun As Boolean, _
       sScaleXLabel As String, _
       sScaleYLabel As String, _
       lChartHeight As Long, _
       lChartWidth As Long
Public AxisLine As Variant, _
       majorticks As Long, _
       minorticks As Long, _
       labelpos As Long, _
       font As String, _
       font_style As String, _
       font_size As Long
Public s1Label As String, _
       s1NewLabel As String, _
       s1Color As Variant, _
       s1Weight As Long, _
       s1Line As Variant, _
       s1PlotOrder As Long
Public gridColor As Variant, _
       gridWeight As Long, _
       gridLine As Variant
Public series_oldstart As String, _
       series_oldend As String, _
       series_start As String, _
       series_end As String, _
       series_sheet As String

Sub run_series_range()

    seriesrows = charts_get_series_rows_range()
    break = InStr(1, seriesrows, "-", vbTextCompare)
    series_oldstart = Mid(seriesrows, 1, break - 1)
    series_start = series_oldstart
    series_oldend = Mid(seriesrows, break + 1)
    series_end = series_oldend
    
    SeriesRange.Show
    If bRun Then
        If StrComp(series_oldstart, series_start, vbTextCompare) <> 0 Or _
           StrComp(series_oldend, series_end, vbTextCompare) <> 0 Then
            Call charts_edit_series_rows_range(series_oldstart, series_oldend, _
                                               series_start, series_end)
        End If
    End If
    
    Call run_Xaxis
    
End Sub

Sub run_Yaxis()
    
    YaxisScale.Show
    If bRun Then
        Call charts_edit_Yaxis_scale(dMin, dMax, dMajor, dMinor, bYreverse)
    End If

    AxisProperties.Show
    If bRun Then
        Call charts_edit_Yaxis_properties(AxisLine, _
                                          majorticks, minorticks, _
                                          labelpos, _
                                          font, font_style, font_size)
    End If
    
    Call run_scale

End Sub

Sub run_Xaxis()
    
    XaxisScale.Show
    If bRun Then
        Call charts_edit_Xaxis_scale(lTickMarkSp, lTickLabelSp, dCrossesAt, _
                                     bYBetween, bXreverse)
    End If
    
    AxisProperties.Show
    If bRun Then
        Call charts_edit_Xaxis_properties(AxisLine, _
                                          majorticks, minorticks, _
                                          labelpos, _
                                          font, font_style, font_size)
    End If
    
    Call run_scale
    
End Sub

Sub run_scale()

    ChartsScale.Show
    If bRun Then
        Call charts_resize_and_format(lChartHeight, lChartWidth)
        Call charts_scale(lScaleXOrigin, lScaleYOrigin, _
                          dScaleXk, dScaleYk, sScaleXLabel, sScaleYLabel)
    End If
End Sub

Sub run_series_properties()
    
    SeriesProperties.Show
    If bRun Then
        Call charts_edit_series_properties(s1Label, s1NewLabel, s1Color, s1Weight, s1Line, s1PlotOrder)
        Call charts_edit_series_legend_properties(s1Label, s1NewLabel, s1Color, s1Weight, s1Line)
    End If
End Sub

Sub run_series_add()
    
    SeriesAdd.Show
    If bRun Then
        Call charts_series_add(series_sheet)
    End If
End Sub

Sub run_series_del()
    
    SeriesDel.Show
    If bRun Then
        Call charts_series_del(series_sheet)
    End If
End Sub

Sub charts_edit_Xaxis_scale(Optional lTickMarkSp As Long = 40, _
                            Optional lTickLabelSp As Long = 80, _
                            Optional dCrossAt As Double = 80, _
                            Optional bYaxisbetweencat As Boolean = False, _
                            Optional bXreverse As Boolean = False)

    n = 0
    
    Dim ch As ChartObject
    For Each ch In ActiveSheet.ChartObjects
        
        If ch.Chart.HasAxis(xlCategory) Then
            
            ch.Chart.axes(xlCategory).CategoryType = xlCategoryScale
            
            With ch.Chart.axes(xlCategory)
                .CrossesAt = dCrossAt
                .TickLabelSpacing = lTickLabelSp
                .TickMarkSpacing = lTickMarkSp
                .AxisBetweenCategories = bYaxisbetweencat
                .ReversePlotOrder = bXreverse
                
            End With
        End If
        n = n + 1
    Next ch
    Call charts_processed_report(n, "Xaxis Scale")
End Sub

Sub charts_edit_Xaxis_properties(Optional AxisLine As Variant = xlAutomatic, _
                                 Optional majorticks As Long = xlTickMarkCross, _
                                 Optional minorticks As Long = xlTickMarkNone, _
                                 Optional labelpos As Long = xlTickLabelPositionNone, _
                                 Optional font As String = "Times New Roman", _
                                 Optional font_style As String = "Regular", _
                                 Optional font_size As Long = 10)
    
    
    n = 0
    
    Dim ch As ChartObject
    For Each ch In ActiveSheet.ChartObjects
        
        ch.Chart.HasAxis(xlCategory) = True
        
        ch.Chart.axes(xlCategory).CategoryType = xlCategoryScale
        
        With ch.Chart.axes(xlCategory)
            
            .HasMajorGridlines = False
            .HasMinorGridlines = False
            
            .Border.Weight = xlHairline
            .Border.LineStyle = AxisLine
            
            .MajorTickMark = majorticks
            .MinorTickMark = minorticks
            .TickLabelPosition = labelpos
            
            .TickLabels.AutoScaleFont = False
            With .TickLabels.font
                .Name = font
                .FontStyle = font_style
                .Size = font_size
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = xlUnderlineStyleNone
                .ColorIndex = xlAutomatic
                .Background = xlAutomatic
            End With
            .TickLabels.NumberFormat = "General"
            .TickLabels.Orientation = xlHorizontal
        End With
        
        n = n + 1
    Next ch
    Call charts_processed_report(n, "Xaxis Properties")
End Sub

Sub charts_edit_Yaxis_scale(Optional dMin As Double = -10#, _
                            Optional dMax As Double = 10#, _
                            Optional dMajorUnit As Double = 1#, _
                            Optional dMinorUnit As Double = 0.5, _
                            Optional bYreverse As Boolean = True)
    
    n = 0
    
    Dim ch As ChartObject
    For Each ch In ActiveSheet.ChartObjects
        
        With ch.Chart.axes(xlValue)
            .MinimumScale = dMin
            .MaximumScale = dMax
            .MajorUnit = dMajorUnit
            .MinorUnit = dMinorUnit
            .ReversePlotOrder = bYreverse
            
            .ScaleType = xlLinear
            .Crosses = xlAxisCrossesAutomatic 'xlMaximum
        
        End With
        n = n + 1
    Next ch
    Call charts_processed_report(n, "Yaxis Scale")
End Sub

Sub charts_edit_Yaxis_properties(Optional AxisLine As Variant = xlAutomatic, _
                                 Optional majorticks As Long = xlTickMarkCross, _
                                 Optional minorticks As Long = xlTickMarkNone, _
                                 Optional labelpos As Long = xlTickLabelPositionNone, _
                                 Optional font As String = "Times New Roman", _
                                 Optional font_style As String = "Regular", _
                                 Optional font_size As Long = 10)
    
    Dim ch As ChartObject, n As Integer
    
    n = 0
    For Each ch In ActiveSheet.ChartObjects
        
        With ch.Chart.axes(xlValue)
            
            .Border.Weight = xlHairline
            .Border.LineStyle = AxisLine
            
            .MajorTickMark = majorticks
            .MinorTickMark = minorticks
            .TickLabelPosition = labelpos
            
            .TickLabels.AutoScaleFont = False
            With .TickLabels.font
                .Name = font
                .FontStyle = font_style
                .Size = font_size
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = xlUnderlineStyleNone
                .ColorIndex = xlAutomatic
                .Background = xlAutomatic
            End With
            .TickLabels.NumberFormat = "General"
            .TickLabels.Orientation = xlHorizontal
        End With
        n = n + 1
    Next ch
    Call charts_processed_report(n, "Yaxis Properties")
End Sub

Sub charts_resize_and_format(Optional h As Long = 80, Optional w As Long = 95)
'
' Warning: reduce runtime by minimising excel window first!
    
    Lindent = 0
    
    goty = False
    gotx = False
    
    Dim ch As ChartObject, n As Integer
    n = 0
    For Each ch In ActiveSheet.ChartObjects
        
        With ch
            
            .Height = h
            .Width = w
            
            If .Chart.HasTitle = True Then
                .Name = "Chart " & .Chart.ChartTitle.Characters.Text
                With .Chart.ChartTitle
                    .Top = 0
                    Tindent = .font.Size + 5
                    
                    'Format Chart Title
                    .Border.Weight = xlHairline
                    .Border.LineStyle = xlNone
                    .Shadow = False
                    .Interior.ColorIndex = xlNone
                End With
            Else
                Tindent = 0
            End If
            
            'Set Plot Area Size (depends on title)
            .Chart.PlotArea.Top = Tindent
            .Chart.PlotArea.Left = Lindent
            .Chart.PlotArea.Height = .Height - Tindent
            .Chart.PlotArea.Width = .Width - Lindent
            
            
            'Format plot area
            .Chart.PlotArea.Interior.ColorIndex = xlNone
            
            'Format chart area
            .Placement = xlFreeFloating 'xlMoveAndSize, xlMove
            .Chart.ChartArea.Border.Weight = 1
            .Chart.ChartArea.Border.LineStyle = 0
            .Chart.ChartArea.Interior.ColorIndex = xlNone 'xlAutomatic
            ActiveSheet.DrawingObjects(.Name).RoundedCorners = False
            ActiveSheet.DrawingObjects(.Name).Shadow = False
            
            Call charts_resize_title(ch)

        End With
        n = n + 1
    Next ch
    
    Call charts_processed_report(n, "Charts Size/Format")
End Sub

Sub charts_resize_title(ch)
    
    If ch.Chart.HasTitle = True Then
        
        With ch
        
            With .Chart.ChartTitle
                .HorizontalAlignment = xlHAlignCenter 'xlLeft
                .VerticalAlignment = xlVAlignCenter 'xlTop
                .Orientation = xlHorizontal
            End With
            .Chart.ChartTitle.AutoScaleFont = False
            With .Chart.ChartTitle.font
                .Name = "Times New Roman"
                .FontStyle = "Regular"
                .Size = 14
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = xlUnderlineStyleNone
                .ColorIndex = xlAutomatic
                .Background = xlAutomatic
            End With
            
            chTitleWidth = .Chart.ChartTitle.Characters.Count
            chTitleSize = .Chart.ChartTitle.font.Size
            chMiddle = (ch.Width / 2) - ((chTitleWidth / 2) * chTitleSize)
            
            .Chart.ChartTitle.Top = 1
            .Chart.ChartTitle.Left = chMiddle
            'Top right
            'shwidth = .Chart.ChartTitle.Width
            '.Chart.ChartTitle.left = chwidth - (shwidth + 5)
            
        End With
    End If
End Sub

Sub charts_edit_series_properties(Optional seriesLabel As String = "Series1", _
                                  Optional seriesNewLabel As String = "Series1", _
                                  Optional seriesColor As Variant = 3, _
                                  Optional seriesWeight As Long = xlHairline, _
                                  Optional seriesLine As Variant = xlContinuous, _
                                  Optional seriesPlotOrder As Long)
    
    Dim ch As ChartObject, n As Integer
    
    For Each ch In ActiveSheet.ChartObjects
        
        With ch
            
            For Each Series In .Chart.SeriesCollection
                
                If StrComp(Series.Name, seriesLabel, vbTextCompare) = 0 Then
                    
                    'Change the series name?
                    If StrComp(seriesLabel, seriesNewLabel, vbTextCompare) <> 0 Then
                        Series.Name = seriesNewLabel
                    End If
                    
                    'Define Series Properties
                    '------------------------
                    sColor = seriesColor
                    sWeight = seriesWeight
                    sLine = seriesLine
                    
                    
                    'Set Series Properties
                    '---------------------
                    With .Chart.SeriesCollection(Series.Name).Border
                        .Color = sColor
                        .Weight = sWeight
                        .LineStyle = sLine
                    End With
                    With .Chart.SeriesCollection(Series.Name)
                        
                        '.PlotOrder = seriesPlotOrder
                        
                        .MarkerBackgroundColorIndex = xlNone
                        .MarkerForegroundColorIndex = xlNone
                        .MarkerStyle = xlMarkerStyleNone
                        .Smooth = True
                        .MarkerSize = 5
                        .Shadow = False
                    End With
                    
                    n = n + 1
                    
                End If
                
            Next Series
            
            ' ------------------------------
            ' Reorder the series plot order
            seriesN = 0
            For Each Series In .Chart.SeriesCollection
                    seriesN = seriesN + 1
                    With .Chart.SeriesCollection(Series.Name)
                        .PlotOrder = seriesN
                    End With
            Next Series
            
        End With
        
    Next ch
    
    Call charts_processed_report(n, "Series Properties")
    
End Sub

Sub charts_edit_series_legend_properties(Optional s1Label As String = "Series1", _
                                         Optional s1NewLabel As String = "Series1", _
                                         Optional s1Color As Variant = vbBlack, _
                                         Optional s1Weight As Long = xlHairline, _
                                         Optional s1Line As Variant = xlContinuous)
    
    xsize = 75
    XOrigin = 100
    YOrigin = 150
    
    'Check if series legend exists
    '-----------------------------
    gotLegend = 0
    For Each sLegend In ActiveSheet.Shapes
        If sLegend.Type = 17 Then
            If StrComp(sLegend.Name, "Legend " & s1Label, vbTextCompare) = 0 Then
                gotLegend = 1
                GoTo ProcessLegend
            End If
        End If
    Next sLegend
    
ProcessLegend:
    
    'Modify or Create Series Legend
    '------------------------------
    If gotLegend Then
        
        ' Modify the legend
        sLegend = "Legend " & s1Label
        ActiveSheet.Shapes(sLegend).Select
        Selection.Characters.Text = s1NewLabel
        Selection.Name = "Legend " & s1NewLabel
        With Selection.Characters.font
            .Name = "Times New Roman"
            .FontStyle = "Regular"
            .Size = 22
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
        End With
        
    Else
        
        ' Create the legend
        LegendName = "Legend " & s1NewLabel
        With ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, XOrigin, YOrigin, xsize, 22)
            .Name = LegendName
        End With
        sLegend = "Legend " & s1NewLabel
        ActiveSheet.Shapes(sLegend).Select
        Selection.Characters.Text = s1NewLabel
        Selection.Name = "Legend " & s1NewLabel
        With Selection.Characters.font
            .Name = "Times New Roman"
            .FontStyle = "Regular"
            .Size = 22
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
        End With
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlTop
            .Orientation = xlHorizontal
            .AutoSize = True
        End With
        Selection.ShapeRange.Fill.Visible = msoFalse
        Selection.ShapeRange.Fill.Transparency = 0#
        Selection.ShapeRange.Line.Weight = 0.75
        Selection.ShapeRange.Line.DashStyle = msoLineSolid
        Selection.ShapeRange.Line.Style = msoLineSingle
        Selection.ShapeRange.Line.Transparency = 0#
        Selection.ShapeRange.Line.Visible = msoFalse
        
    End If
    
    
    'Check if series line exists
    '---------------------------
    gotLine = 0
    For Each sLine In ActiveSheet.Shapes
        If sLine.Type = msoLine Then
            If StrComp(sLine.Name, "Line " & s1Label, vbTextCompare) = 0 Then
                gotLine = 1
                GoTo ProcessLine
            End If
        End If
    Next sLine
    
ProcessLine:
    
    
    'Can be one of the following MsoLineDashStyle constants: msoLineDash, msoLineDashDot,
    'msoLineDashDotDot, msoLineDashStyleMixed, msoLineLongDash, msoLineLongDashDot,
    'msoLineRoundDot, msoLineSolid, or msoLineSquareDot. Read/write Long.
    
    'Can be one of the following XlLineStyle constants: xlContinuous, xlDash,
    'xlDashDot, xlDashDotDot, xlDot, xlDouble, xlSlantDashDot, or
    'xlLineStyleNone. Read/write Variant.
    
    ' Change s1Line to an msoLineDashStyle constant
    Dim LegendLineDashStyle As Long
    Select Case s1Line
        Case xlLineStyleNone
            LegendLineDashStyle = msoLineSolid
            s1Color = RGB(255, 255, 255) ' White
        Case xlAutomatic
            LegendLineDashStyle = msoLineSolid
        Case xlContinuous
            LegendLineDashStyle = msoLineSolid
        Case xlDash
            LegendLineDashStyle = msoLineDash
        Case xlDashDot
            LegendLineDashStyle = msoLineDashDot
        Case xlDashDotDot
            LegendLineDashStyle = msoLineDashDotDot
        Case xlDot
            LegendLineDashStyle = msoLineRoundDot
        Case xlDouble
            LegendLineDashStyle = msoLineSolid
        Case xlSlantDashDot
            LegendLineDashStyle = msoLineLongDashDot
    End Select
    
    
    'Modify or Create Series Line
    '----------------------------
    If gotLine Then
        
        ' Modify the line
        With sLine
            .Name = "Line " & s1NewLabel
            .Line.Weight = s1Weight
            .Line.DashStyle = LegendLineDashStyle
            .Line.Style = msoLineSingle
            
            '.Line.ForeColor.SchemeColor = 10 ' Red
            .Line.ForeColor.RGB = s1Color
            
            .Line.BackColor.RGB = RGB(255, 255, 255) ' White
            .Line.Visible = msoTrue
            .Line.Transparency = 0#
            .Line.BeginArrowheadStyle = msoArrowheadNone
            .Line.BeginArrowheadLength = msoArrowheadLengthMedium
            .Line.BeginArrowheadWidth = msoArrowheadWidthMedium
            .Line.EndArrowheadStyle = msoArrowheadNone
            .Line.EndArrowheadLength = msoArrowheadLengthMedium
            .Line.EndArrowheadWidth = msoArrowheadWidthMedium
        End With
        
    Else
        
        ' Create the line
        With ActiveSheet.Shapes.AddLine(XOrigin, YOrigin, XOrigin + xsize, YOrigin)
            .Name = "Line " & s1NewLabel
            .Line.Weight = s1Weight
            .Line.DashStyle = LegendLineDashStyle
            .Line.Style = msoLineSingle
            
            '.Line.ForeColor.SchemeColor = 10 ' Red
            .Line.ForeColor.RGB = s1Color
            
            .Line.BackColor.RGB = RGB(255, 255, 255) ' White
            .Line.Visible = msoTrue
            .Line.Transparency = 0#
            .Line.BeginArrowheadStyle = msoArrowheadNone
            .Line.BeginArrowheadLength = msoArrowheadLengthMedium
            .Line.BeginArrowheadWidth = msoArrowheadWidthMedium
            .Line.EndArrowheadStyle = msoArrowheadNone
            .Line.EndArrowheadLength = msoArrowheadLengthMedium
            .Line.EndArrowheadWidth = msoArrowheadWidthMedium
        End With
    End If
    
End Sub
Function charts_get_series_rows_range() As String

    Dim ch As ChartObject, _
        formula As String, _
        old_start_row As String, _
        old_end_row As String
        
    'Get current start and end rows
    For Each ch In ActiveSheet.ChartObjects
        With ch
            
            formula = .Chart.SeriesCollection(1).formula
            
            dollar1_pos = InStr(1, formula, "$", vbTextCompare)
            dollar2_pos = InStr(dollar1_pos + 1, formula, "$", vbTextCompare)
            colon1_pos = InStr(dollar2_pos + 1, formula, ":", vbTextCompare)
            dollar3_pos = InStr(colon1_pos + 1, formula, "$", vbTextCompare)
            dollar4_pos = InStr(dollar3_pos + 1, formula, "$", vbTextCompare)
            comma2_pos = InStr(dollar4_pos + 1, formula, ",", vbTextCompare)
            
            old_start_row = Mid(formula, dollar2_pos + 1, colon1_pos - (dollar2_pos + 1))
            
            old_end_row = Mid(formula, dollar4_pos + 1, comma2_pos - (dollar4_pos + 1))
            
            'MsgBox formula & " " & old_start_str & " " & old_end_str
            
        End With
        GoTo Fin
    Next ch
Fin:
    charts_get_series_rows_range = old_start_row & "-" & old_end_row
    
End Function

Sub charts_series_add(series_sheet As String)

    Dim ch As ChartObject, n As Integer

    n = 0
    For Each ch In ActiveSheet.ChartObjects
        With ch
            
            formula = .Chart.SeriesCollection(1).FormulaR1C1
            
            GoSub REPLACE_SERIES_SHEET
            
            Set ns = .Chart.SeriesCollection.NewSeries
            
            ns.FormulaR1C1 = formula
            ns.Name = series_sheet
            
            'MsgBox formula
            
        End With
        n = n + 1
    Next ch
    
    GoTo Fin
    
REPLACE_SERIES_SHEET:

    ' First find all the commas in the formula
    comma1 = InStr(1, formula, ",", vbTextCompare)
    comma2 = InStr(comma1 + 1, formula, ",", vbTextCompare)
    comma3 = InStr(comma2 + 1, formula, ",", vbTextCompare)
    ' Use the commas to find the start and end of the formula
    formula_start = Mid(formula, 1, comma1)
    formula_end = Mid(formula, comma3)
    ' Use the commas to find the x/y range formulas
    xrange = Mid(formula, comma1 + 1, (comma2 - comma1 - 1))
    yrange = Mid(formula, comma2 + 1, (comma3 - comma2 - 1))
    ' Replace the Yrange sheet reference
    y! = InStr(1, yrange, "!", vbTextCompare)
    yrangeRC = Mid(yrange, y!)
    
    yrange = series_sheet & yrangeRC
    
    formula = formula_start & xrange & "," & yrange & formula_end
Return
    
Fin:
    Call charts_processed_report(n, "Series Add")
End Sub

Sub charts_series_del(series_sheet As String)
    
    Dim ch As ChartObject, n As Integer
    
    n = 0
    For Each ch In ActiveSheet.ChartObjects
        With ch
            ch.Select
            For Each Series In .Chart.SeriesCollection
            
                formula = Series.FormulaR1C1
                
                gotseries = InStr(1, formula, series_sheet, vbTextCompare)
                
                If gotseries > 0 Then
                    GoSub DELETE_SERIES
                End If
                
                'MsgBox formula
                
            Next Series
            
        End With
        n = n + 1
    Next ch
    
    GoTo Fin
    
DELETE_SERIES:

    ' First find all the commas in the formula
    comma1 = InStr(1, formula, ",", vbTextCompare)
    comma2 = InStr(comma1 + 1, formula, ",", vbTextCompare)
    comma3 = InStr(comma2 + 1, formula, ",", vbTextCompare)
    ' Use the commas to find the y range formula
    yrange = Mid(formula, comma2 + 1, (comma3 - comma2 - 1))
    ' Try to match the Yrange sheet reference
    y! = InStr(1, yrange, "!", vbTextCompare)
    yrange_sheet = Mid(yrange, 1, y! - 1)
    
    'formula = yrange_sheet
    
    If StrComp(yrange_sheet, series_sheet, vbTextCompare) = 0 Then
        'formula = yrange_sheet & " - remove this one"
        'Remove this series
        Series.Delete
    End If
    
Return
    
Fin:
    Call charts_processed_report(n, "Series Delete")
End Sub

Sub charts_edit_series_rows_range(old_start_row As String, _
                                  old_end_row As String, _
                                  new_start_row As String, _
                                  new_end_row As String)

    Dim ch As ChartObject, _
        formula As String, _
        old_start_str As String, _
        old_end_str As String, _
        new_start_str As String, _
        new_end_str As String

    
    old_start_str = "$" & old_start_row & ":"
    old_end_str = "$" & old_end_row & ","
    
    new_start_str = "$" & new_start_row & ":"
    new_end_str = "$" & new_end_row & ","
    
    'new_xstart_row = 0      'set to 0 to skip xrange update
    'new_xend_row = 401
    
    n = 0
    For Each ch In ActiveSheet.ChartObjects
        With ch
        
            For Each Series In .Chart.SeriesCollection
            
                formula = .Chart.SeriesCollection(Series.Name).formula
                GoSub REPLACE_SERIES_ROWS
                .Chart.SeriesCollection(Series.Name).formula = formula
                
                'MsgBox formula
                'If new_xstart_row > 0 Then
                '    Xrange = "R" & new_xstart_row & "C126:" & "R" & new_xend_row & "C126"
                '    .Chart.SeriesCollection(1).XValues = ("='" & ActiveWorkbook.Sheets(1) & "'!" & Xrange)
                'End If
            Next Series
        End With
        n = n + 1
    Next ch
   
    GoTo Fin
    
REPLACE_SERIES_ROWS:

    If StrComp(old_start_str, new_start_str, vbTextCompare) <> 0 Then
        While (InStr(formula, old_start_str) > 0)
            begin_pos = InStr(1, formula, old_start_str, vbTextCompare)
            end_pos = begin_pos + Len(old_start_str)
            
            begin_str = Mid(formula, 1, begin_pos - 1)
            end_str = Mid(formula, end_pos)
            
            formula = begin_str & new_start_str & end_str
        Wend
    End If
    
    If StrComp(old_end_str, new_end_str, vbTextCompare) <> 0 Then
        While (InStr(formula, old_end_str) > 0)
        
            begin_pos = InStr(1, formula, old_end_str, vbTextCompare)
            end_pos = begin_pos + Len(old_end_str)
            
            begin_str = Mid(formula, 1, begin_pos - 1)
            end_str = Mid(formula, end_pos)
            
            formula = begin_str & new_end_str & end_str
        Wend
    End If
    Return
    
Fin:
    Call charts_processed_report(n, "Series Range")
End Sub

Sub charts_scale(Optional XOrigin As Long = 1000, _
                 Optional YOrigin As Long = 150, _
                 Optional xk As Double = 1#, _
                 Optional yk As Double = 1#, _
                 Optional x_metric_text = "ms", _
                 Optional y_metric_text = "uV")
    
    ' Get scaling parameters from first chart
    Dim ch As ChartObject
    For Each ch In ActiveSheet.ChartObjects

        With ch
            If .Chart.HasAxis(xlCategory) = True Then
                With .Chart.axes(xlCategory)
                    
                    ' requires definition of msec range in 'special' series
                    cNames = .CategoryNames
                    
                    cfirst = LBound(cNames)
                    clast = UBound(cNames)
                    nCategories = clast - cfirst
                    
                    x_min = cNames(cfirst)
                    x_max = cNames(clast)
                    
                    x_sample_rate = cNames(cfirst + 1) - x_min
                    
                    x_spacing = .TickMarkSpacing
                    x_units = x_sample_rate * x_spacing
                    x_length = .Width
                    x_pixels = x_length / ((x_max - x_min) / x_units)
                    
                End With
            End If
            If .Chart.HasAxis(xlValue) = True Then
                With .Chart.axes(xlValue)
                    y_reverse = .ReversePlotOrder
                    y_units = .MajorUnit
                    y_min = .MinimumScale
                    y_max = .MaximumScale
                    y_length = .Height
                    y_pixels = y_length / ((y_max - y_min) / y_units)
                    goty = True
                End With
            End If
        End With
        GoTo Skip  ' Just get data from first chart
    Next ch
Skip:
    
    ' Define Names for Scale Shapes
    ScaleLines = Array("x_scale_line", "y_scale_line")
    ScaleLabels = Array("x_scale_metric", "x_scale_min", "x_scale_max", "y_scale_metric", "y_scale_min", "y_scale_max")
    
    ' Remove Existing Scale Shapes, if present
    For Each shName In ScaleLines
        On Error Resume Next
        ActiveSheet.Shapes(shName).Delete
    Next shName
    For Each shName In ScaleLabels
        On Error Resume Next
        ActiveSheet.Shapes(shName).Delete
    Next shName
    
    'Multiply the whole scale by x,y constants xk, yk
    'x_units = x_units * xk
    'x_pixels = x_pixels * xk
    'y_units = y_units * yk
    'y_pixels = y_pixels * yk
    
    'Add x scale line
    '----------------
    With ActiveSheet.Shapes.AddLine(XOrigin, YOrigin, XOrigin + x_length, YOrigin)
        .Name = "x_scale_line"
        ' set other line properties here
    End With
    
    'Add x labels
    '------------
    x_label_top = YOrigin + 5
    
    x_min_left = XOrigin - 30
    With ActiveSheet.Shapes.AddLabel(msoTextOrientationHorizontal, x_min_left, x_label_top, 0#, 0#)
        .TextFrame.Characters.Text = CStr(x_min)
        .Name = "x_scale_min"
    End With
    
    x_max_left = XOrigin + x_length - 30
    With ActiveSheet.Shapes.AddLabel(msoTextOrientationHorizontal, x_max_left, x_label_top, 0#, 0#)
        .TextFrame.Characters.Text = CStr(x_max)
        .Name = "x_scale_max"
    End With
    
    x_metric_left = XOrigin + x_length + 10
    x_metric_top = YOrigin - (14 / 2)
    With ActiveSheet.Shapes.AddLabel(msoTextOrientationHorizontal, x_metric_left, x_metric_top, 0#, 0#)
        .TextFrame.Characters.Text = x_metric_text
        .Name = "x_scale_metric"
    End With
    
    
    'Add y scale line
    '----------------
    baseline = x_min * (x_pixels / x_units)
    If x_min < 0 Then
        yx_origin = XOrigin - baseline
    Else
        yx_origin = XOrigin
    End If
    
    If y_reverse Then
        y_topPC = y_min / (y_max - y_min)
        y_topPC = Abs(y_topPC)
        y_bottomPC = y_max / (y_max - y_min)
        y_bottomPC = Abs(y_bottomPC)
        y_top = YOrigin - (y_length * y_topPC)
        y_bottom = YOrigin + (y_length * y_bottomPC)
    Else
        y_topPC = y_max / (y_max - y_min)
        y_topPC = Abs(y_topPC)
        y_bottomPC = y_min / (y_max - y_min)
        y_bottomPC = Abs(y_bottomPC)
        y_top = YOrigin + (y_length * y_topPC)
        y_bottom = YOrigin - (y_length * y_bottomPC)
    End If
    
    With ActiveSheet.Shapes.AddLine(yx_origin, y_top, yx_origin, y_bottom)
        .Name = "y_scale_line"
    End With
    
    'Add y labels
    '--------------------
    y_label_left = yx_origin
    y_metric_top = y_top
    If y_reverse Then
        y_max_top = y_bottom
        y_min_top = y_top - 20
    Else
        y_max_top = y_top - 20
        y_min_top = y_bottom
    End If
    
    With ActiveSheet.Shapes.AddLabel(msoTextOrientationHorizontal, y_label_left, y_min_top, 0#, 0#)
        .TextFrame.Characters.Text = y_min
        .Name = "y_scale_min"
    End With
    With ActiveSheet.Shapes.AddLabel(msoTextOrientationHorizontal, y_label_left, y_max_top, 0#, 0#)
        .TextFrame.Characters.Text = "+" & y_max
        .Name = "y_scale_max"
    End With
    With ActiveSheet.Shapes.AddLabel(msoTextOrientationHorizontal, (y_label_left - 30), y_metric_top, 0#, 0#)
        .TextFrame.Characters.Text = y_metric_text
        .Name = "y_scale_metric"
    End With
    
    
    'format text labels
    '------------------
    For Each shName In ScaleLabels
        With ActiveSheet.Shapes(shName).TextFrame
            .HorizontalAlignment = xlHAlignRight
            .VerticalAlignment = xlCenter
            .AutoMargins = True
            .AutoSize = True
            With .Characters.font
                .Name = "Times New Roman"
                .FontStyle = "Regular"
                .Size = 14
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = xlUnderlineStyleNone
                .ColorIndex = xlAutomatic
            End With
        End With
    Next shName
End Sub

Sub charts_sheet_pagesetup(ch)

    'For Each ch In ActiveWorkbook.Charts
        
        ActiveWindow.Zoom = 75
        
        With ch.PageSetup
            .LeftHeader = ""
            .CenterHeader = ""
            .RightHeader = ""
            .LeftFooter = ""
            .CenterFooter = ""
            .RightFooter = ""
            .LeftMargin = Application.CentimetersToPoints(0)
            .RightMargin = Application.CentimetersToPoints(0)
            .TopMargin = Application.CentimetersToPoints(0)
            .BottomMargin = Application.CentimetersToPoints(0)
            .HeaderMargin = Application.CentimetersToPoints(0)
            .FooterMargin = Application.CentimetersToPoints(0)
            .ChartSize = xlFullPage 'xlFitToPage
            .PrintQuality = 300
            .CenterHorizontally = False
            .CenterVertically = False
            .Orientation = xlLandscape 'xlPortrait
            .Draft = False
            .PaperSize = xlPaperA4
            .FirstPageNumber = xlAutomatic
            .BlackAndWhite = False
            .Zoom = 100
        End With
    'Next ch
    
End Sub

Sub charts_add_gridlines(Optional gridColor As Variant = 15, _
                         Optional gridWeight As String = "xlHairline", _
                         Optional gridLine As String = "xlContinuous")

    With ActiveChart.axes(xlCategory)
        .HasMajorGridlines = True
        .HasMinorGridlines = False
        With .MajorGridlines.Border
            .ColorIndex = gridColor        ' Gray 25%
            .Weight = gridWeight
            .LineStyle = gridLine
        End With
    End With
    
    With ActiveChart.axes(xlValue)
        .HasMajorGridlines = True
        .HasMinorGridlines = False
        With .MajorGridlines.Border
            .ColorIndex = gridColor        ' Gray 25%
            .Weight = gridWeight
            .LineStyle = gridLine
        End With
    End With

End Sub

Sub activechart_resize_chartarea(Optional w = 30, Optional h = 21)
    
    ' w & h are in cm and set to A4 page size by default
    
    If ActiveChart.PageSetup.ChartSize = xlFullPage Then
        ' Can't do anything
    Else
        With ActiveChart
            
            .ChartArea.Width = .Application.CentimetersToPoints(w)
            .ChartArea.Height = .Application.CentimetersToPoints(h)
            
        End With
    End If
End Sub

Sub charts_processed_report(n, Optional title As String = "Charts Processed")
    Dim Msg, Style, Response
    Msg = "Modified " & n & " Charts."
    Style = vbOKOnly + vbInformation + vbDefaultButton3
    Response = MsgBox(Msg, Style, title)
End Sub
