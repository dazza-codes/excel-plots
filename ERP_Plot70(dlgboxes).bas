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
       sScaleYLabel As String
Public AxisLine As Variant, _
       majorticks As Long, _
       minorticks As Long, _
       labelpos As Long, _
       font As String, _
       font_style As String, _
       font_size As Long
Public s1Color As Variant, _
       s2Color As Variant, _
       s1Weight As Long, _
       s2Weight As Long, _
       s1Line As Variant, _
       s2Line As Variant
Public gridColor As Variant, _
       gridWeight As Long, _
       gridLine As Variant
Public series_oldstart As String, _
       series_oldend As String, _
       series_start As String, _
       series_end As String

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

    ChartsScaleLegend.Show
    If bRun Then
        Call charts_resize_and_format
        Call charts_scale(lScaleXOrigin, lScaleYOrigin, _
                          dScaleXk, dScaleYk, sScaleXLabel, sScaleYLabel)
    End If
End Sub

Sub run_series_properties()
    
    SeriesProperties.Show
    If bRun Then
        Call charts_edit_series_properties(s1Color, s2Color, _
                                           s1Weight, s2Weight, _
                                           s1Line, s2Line)
        Call charts_edit_series_legend_properties(s1Color, s2Color)
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
            
            ch.Chart.Axes(xlCategory).CategoryType = xlCategoryScale
            
            With ch.Chart.Axes(xlCategory)
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
        
        ch.Chart.Axes(xlCategory).CategoryType = xlCategoryScale
        
        With ch.Chart.Axes(xlCategory)
            
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
        
        With ch.Chart.Axes(xlValue)
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
        
        With ch.Chart.Axes(xlValue)
            
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

Sub charts_resize_and_format(Optional h = 80, Optional w = 95)
'
' Warning: reduce runtime by minimising excel window first!
    
    Lindent = 0
    
    goty = False
    gotx = False
    
    Dim ch As ChartObject, n As Integer
    n = 0
    For Each ch In ActiveSheet.ChartObjects
        
        With ch
        
            If .Chart.HasTitle = True Then
                .Name = "Chart " & .Chart.ChartTitle.Characters.text
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
            
            .Height = h + Tindent + 5
            .Width = w + Lindent + 5
            
            'Set Plot Area Size (depends on title)
            .Chart.PlotArea.Top = Tindent
            .Chart.PlotArea.left = Lindent
            .Chart.PlotArea.Height = h
            .Chart.PlotArea.Width = w
            
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
            .Chart.ChartTitle.left = chMiddle
            'Top right
            'shwidth = .Chart.ChartTitle.Width
            '.Chart.ChartTitle.left = chwidth - (shwidth + 5)
            
        End With
    End If
End Sub

Sub charts_edit_series_properties(Optional s1Color As Variant = 3, _
                                  Optional s2Color As Variant = 5, _
                                  Optional s1Weight As Long = xlHairline, _
                                  Optional s2Weight As Long = xlHairline, _
                                  Optional s1Line As Variant = xlContinuous, _
                                  Optional s2Line As Variant = xlContinuous)

    Dim ch As ChartObject
    
    For Each ch In ActiveSheet.ChartObjects
        
        With ch
        
            For Each Series In .Chart.SeriesCollection
        
                
                'Define Series Properties
                '------------------------
                Select Case Series.Name
                    Case "Series1"
                        sColor = s1Color
                        sWeight = s1Weight
                        sLine = s1Line
                        sPlotOrder = 2
                    Case "Series2"
                        sColor = s2Color
                        sWeight = s2Weight
                        sLine = s2Line
                        sPlotOrder = 1
                End Select
                
                'Set Series Properties
                '---------------------
                With .Chart.SeriesCollection(Series.Name).Border
                    .Color = sColor
                    .Weight = sWeight
                    .LineStyle = sLine
                End With
                With .Chart.SeriesCollection(Series.Name)
                    
                    .PlotOrder = sPlotOrder
                    
                    .MarkerBackgroundColorIndex = xlNone
                    .MarkerForegroundColorIndex = xlNone
                    .MarkerStyle = xlMarkerStyleNone
                    .Smooth = False
                    .MarkerSize = 5
                    .Shadow = False
                End With
        
            Next Series
            
        End With
        
    Next ch
    
End Sub

Sub charts_edit_series_legend_properties(Optional s1Color As Variant = 3, _
                                         Optional s2Color As Variant = 5)

    xsize = 75
    XOrigin = 175
    YOrigin = 150

    SeriesLines = Array("Line Series1", "Line Series2")
    'Remove current series lines
    '--------------------------
    For Each sName In SeriesLines
        On Error Resume Next
        ActiveSheet.Shapes(sName).Delete
    Next sName
    
    'Create New Horizontal Series Lines
    '----------------------------------
    With ActiveSheet.Shapes.AddLine(XOrigin, YOrigin, XOrigin + xsize, YOrigin)
        .Name = SeriesLines(0)
        .Line.Weight = 5#
        .Line.DashStyle = msoLineSolid
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
    
    With ActiveSheet.Shapes.AddLine(XOrigin, YOrigin + 25, XOrigin + xsize, YOrigin + 25)
        .Name = SeriesLines(1)
        .Line.Weight = 5#
        .Line.DashStyle = msoLineSolid
        .Line.Style = msoLineSingle
        
        .Line.ForeColor.RGB = s2Color
        
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
                 Optional xk As Double = 5#, _
                 Optional yk As Double = 5#, _
                 Optional x_scale_text = " ms", _
                 Optional y_scale_text = "uV")
    
    ' Get scaling parameters from first chart
    Dim ch As ChartObject
    For Each ch In ActiveSheet.ChartObjects

        With ch
            If .Chart.HasAxis(xlCategory) = True Then
                With .Chart.Axes(xlCategory)
                    
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
                With .Chart.Axes(xlValue)
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
    ScaleLabels = Array("x_scale_label", "y_scale_label", "py_scale_label", "ny_scale_label")
    
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
    x_units = x_units * xk
    x_pixels = x_pixels * xk
    y_units = y_units * yk
    y_pixels = y_pixels * yk
    
    'Add x scale line
    '----------------
    With ActiveSheet.Shapes.AddLine(XOrigin, YOrigin, XOrigin + x_pixels, YOrigin)
        .Name = "x_scale_line"
        ' set other line properties here
    End With
    
    'Add x label
    '-----------
    x_label_left = XOrigin + 20
    x_label_top = YOrigin + 5
    x_scale_text = CStr(x_min) & " - " & CStr(x_max) & x_scale_text
    
    With ActiveSheet.Shapes.AddLabel(msoTextOrientationHorizontal, x_label_left, x_label_top, 0#, 0#)
        .Name = "x_scale_label"
        .TextFrame.Characters.text = x_scale_text
    End With
        
    'Add y scale line
    '----------------
    baseline = x_min * (x_pixels / x_units)
    If x_min < 0 Then
        yx_origin = XOrigin - baseline
    Else
        yx_origin = XOrigin
    End If
    y_top = YOrigin + y_pixels
    y_bottom = YOrigin - y_pixels
    
    With ActiveSheet.Shapes.AddLine(yx_origin, y_top, yx_origin, y_bottom)
        .Name = "y_scale_line"
    End With
    
    'Add amplitude labels
    '--------------------
    y_label_left = XOrigin - 50
    y_label_top = y_bottom
    
    With ActiveSheet.Shapes.AddLabel(msoTextOrientationHorizontal, y_label_left, y_label_top, 0#, 0#)
        .TextFrame.Characters.text = y_scale_text
        .Name = "y_scale_label"
    End With
    
    If y_reverse Then
        py_label_top = y_top
        ny_label_top = y_bottom - 20
    Else
        py_label_top = y_bottom - 20
        ny_label_top = y_top
    End If
    
    py_label_left = XOrigin
    py_scale_text = "+" & CStr(y_units)
    With ActiveSheet.Shapes.AddLabel(msoTextOrientationHorizontal, py_label_left, py_label_top, 0#, 0#)
        .TextFrame.Characters.text = py_scale_text
        .Name = "py_scale_label"
    End With
    
    ny_label_left = XOrigin
    ny_scale_text = "-" & CStr(y_units)
    With ActiveSheet.Shapes.AddLabel(msoTextOrientationHorizontal, ny_label_left, ny_label_top, 0#, 0#)
        .TextFrame.Characters.text = ny_scale_text
        .Name = "ny_scale_label"
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

    With ActiveChart.Axes(xlCategory)
        .HasMajorGridlines = True
        .HasMinorGridlines = False
        With .MajorGridlines.Border
            .ColorIndex = gridColor        ' Gray 25%
            .Weight = gridWeight
            .LineStyle = gridLine
        End With
    End With
    
    With ActiveChart.Axes(xlValue)
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
