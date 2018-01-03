Attribute VB_Name = "ERP_Plot124"
Sub charts_processed_report(n)

    Dim Msg, Style, Title, Response
    Title = "Number of Charts"
    Msg = "Modified " & n & " Charts."
    Style = vbOKOnly + vbInformation + vbDefaultButton3
    Response = MsgBox(Msg, Style, Title)

End Sub

Sub ptsd_pet_plots_all()

    Dim plotPath, plotBook, plotSheet, dataPath, dataSheet As String
    
    group = Array("c", "p")
    groupNames = Array("Control", "PTSD")
    groupFileNames = Array("cont", "ptsd")
    
    comp = Array("sa", "wm", "ea", "dt")
    
    Dim beginTime, points As Integer, SampleRate As Double
    
    beginTime = -200
    SampleRate = 2.5
    points = 681
    
    xaxis_title = "msec"
    
    data = Array("volt", "scd")
    
    For Each dat In data
        
        If dat = "volt" Then
            plotPath = "D:\MyDocuments\THESIS\results\erp plots\All - Voltages\"
            dataPath = "D:\MyDocuments\THESIS\results\erp plots\All - Voltages\text files\"
            file_ext = "_link14hz"
            title_ext = ", Voltage (14 Hz lowpass)"
            yaxis_title = "uV"
        Else
            plotPath = "D:\MyDocuments\THESIS\results\erp plots\All - SCD\"
            dataPath = "D:\MyDocuments\THESIS\results\erp plots\All - SCD\text files\"
            file_ext = "_scd14hz"
            title_ext = ", Scalp Current Density (14 Hz lowpass)"
            yaxis_title = "uA/m^3"
        End If
        
        'Call ptsd_pet_plots_124dif(plotPath, plotBook, plotSheet, dataPath, dataSheet, _
                                   group, groupNames, groupFileNames, _
                                   comp, _
                                   beginTime, SampleRate, _
                                   file_ext, _
                                   title_ext, xaxis_title, yaxis_title)
        
        Call ptsd_pet_plots_124overlay(plotPath, plotBook, plotSheet, dataPath, dataSheet, _
                                   group, groupNames, groupFileNames, _
                                   comp, _
                                   beginTime, SampleRate, _
                                   file_ext, _
                                   title_ext, xaxis_title, yaxis_title)
    
        'Call ptsd_pet_plots_regions(plotPath, plotBook, plotSheet, dataPath, dataSheet, _
                                   group, groupNames, groupFileNames, _
                                   comp, _
                                   beginTime, SampleRate, points, _
                                   file_ext, _
                                   title_ext, xaxis_title, yaxis_title)
                                   
    Next dat

End Sub
Sub ptsd_pet_plots_124dif(plotPath, plotBook, plotSheet, dataPath, dataSheet, _
                               group, groupNames, groupFileNames, _
                               comp, _
                               beginTime, SampleRate, _
                               file_ext, _
                               title_ext, xaxis_title, yaxis_title)
'
' Macro coded 30/03/00 by Darren Weber
'

'
    
    For cp = 0 To UBound(comp)
        For g = 0 To UBound(group)
                        
            plotBook = (comp(cp) & file_ext & "_" & groupFileNames(g) & "_124dif.xls")
    
            Workbooks.Add
            ActiveWorkbook.SaveAs FileName:=(plotPath & plotBook), _
                FileFormat:=xlNormal, Password:="", WriteResPassword:="", _
                ReadOnlyRecommended:=False, CreateBackup:=False
            Cells(1, 1).Select
            
            If (comp(cp) = "sa") Then
                cond = Array("ouc", "oac")
                condName = Array("Fixed Unattended Word", "Fixed Attended Word")
            ElseIf (comp(cp) = "wm") Then
                cond = Array("oac", "tac")
                condName = Array("Fixed Attended Word", "Variable Attended Word")
            ElseIf (comp(cp) = "ea") Then
                cond = Array("oac", "oat")
                condName = Array("Fixed Attended Common", "Fixed Attended Target")
            ElseIf (comp(cp) = "dt") Then
                cond = Array("tuc", "tud")
                condName = Array("Variable Unattended Common", "Variable Unattended Distractor")
            End If
            
            For C = 0 To UBound(condName)
                
                ' Open text data file and copy data
                dataSheet = (group(g) & cond(C) & file_ext & ".xls")
                Workbooks.Open FileName:=(dataPath & dataSheet)
                Range(Cells(1, 1), Cells(680, 124)).Select
                Selection.Copy
                
                ' Add new sheet and paste selection
                difdataSheet = (group(g) & cond(C) & file_ext)
                Windows(plotBook).Activate
                Worksheets.Add.Name = difdataSheet
                ActiveSheet.PASTE
                
                ' Save and close data file
                Workbooks(dataSheet).Save
                Workbooks(dataSheet).Close
                
            Next C
            
            ' Add difference sheet and calculate difference values
            
            ASheet = (group(g) & cond(1) & file_ext)
            BSheet = (group(g) & cond(0) & file_ext)
            
            difdataSheet = (group(g) & cond(1) & " - " & group(g) & cond(0) & file_ext)
            Worksheets.Add.Name = difdataSheet

            Cells(1, 1).Activate
            ActiveCell.FormulaR1C1 = ("=" & ASheet & "!RC - " & BSheet & "!RC")
            Range(Cells(1, 1), Cells(680, 1)).Select
            Selection.FillDown
            Range(Cells(1, 1), Cells(680, 124)).Select
            Selection.FillRight
            Selection.Copy
            Selection.PasteSpecial PASTE:=xlValues, Operation:=xlNone, SkipBlanks:= _
                False, Transpose:=False
            Cells(1, 1).Select
            
            ' Remove individual condition sheets
            'Worksheets(ASheet).Delete
            'Worksheets(BSheet).Delete
            
            ' Copy last data row into next row
            Worksheets(difdataSheet).Select
            Cells(681, 1).Activate
            ActiveCell.FormulaR1C1 = "=R[-1]C"
            Range(Cells(681, 1), Cells(681, 124)).Select
            Selection.FillRight
            
            ' Calculate column & overall min/max
            Cells(683, 1).Activate
            ActiveCell.FormulaR1C1 = "=MIN(R1C:R681C)"
            Cells(684, 1).Activate
            ActiveCell.FormulaR1C1 = "=MAX(R1C:R681C)"
            Range(Cells(683, 1), Cells(684, 124)).Select
            Selection.FillRight
            
            Cells(683, 126).Activate
            ActiveCell.FormulaR1C1 = "=MIN(RC1:RC124)"
            Cells(684, 126).Activate
            ActiveCell.FormulaR1C1 = "=MAX(RC1:RC124)"
            
            ' Add timescale data to column 126 (using formula and paste values)
            Cells(1, 126).Activate
            ActiveCell.FormulaR1C1 = beginTime
            Cells(2, 126).Activate
            ActiveCell.FormulaR1C1 = ("=R[-1]C + " & SampleRate)
            Cells(2, 126).Select
            Selection.AutoFill Destination:=Range(Cells(2, 126), Cells(681, 126)), Type:=xlFillDefault
            Range(Cells(2, 126), Cells(681, 126)).Select
            Selection.Copy
            Selection.PasteSpecial PASTE:=xlValues, Operation:=xlNone, SkipBlanks:= _
                False, Transpose:=False
            Cells(1, 1).Select
            
            ' Add Chart
            Dim chartRange As Range
            Range(Cells(1, 1), Cells(681, 124)).Select
            Set chartRange = Range(Cells(1, 1), Cells(681, 124))
            Charts.Add
            ActiveChart.ChartType = xlLine
            ActiveChart.SetSourceData Source:=chartRange, _
                PlotBy:=xlColumns
            
            For s = 1 To 124
                ActiveChart.SeriesCollection(s).XValues = ("='" & difdataSheet & "'!" & "R1C126:R681C126")
            Next s
            
            ActiveChart.Location Where:=xlLocationAsNewSheet, Name:=(difdataSheet & " plot")
            
            ' Chart Title Definitions
            With ActiveChart
                .HasTitle = True
                .ChartTitle.Characters.Text = (groupNames(g) & title_ext & Chr$(10) & condName(1) & " - " & condName(0))
                .Axes(xlCategory, xlPrimary).HasTitle = True
                .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = xaxis_title
                .Axes(xlValue, xlPrimary).HasTitle = True
                .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = yaxis_title
            End With
            
            ' Remove Legend
            ActiveChart.HasLegend = False
            
            ' Chart Axis Formats
            With ActiveChart
                .HasAxis(xlCategory, xlPrimary) = True
                .HasAxis(xlValue, xlPrimary) = True
            End With
            
            ActiveChart.Axes(xlCategory, xlPrimary).CategoryType = xlCategoryScale
            Call charts_add_gridlines
                                        
            ActiveChart.Axes(xlCategory).Select
            With Selection.Border
                .Weight = xlHairline
                .LineStyle = xlAutomatic
            End With
            With Selection
                .MajorTickMark = xlCross
                .MinorTickMark = xlOutside
                .TickLabelPosition = xlNextToAxis
            End With
            With ActiveChart.Axes(xlCategory)
                .crossesat = 1
                .ticklabelspacing = 40
                .tickmarkspacing = 40
                .AxisBetweenCategories = False
                .ReversePlotOrder = False
            End With
            ActiveChart.Axes(xlValue).Select
            With ActiveChart.Axes(xlValue)
                .MinimumScaleIsAuto = True
                .MaximumScaleIsAuto = True
                .MinorUnitIsAuto = True
                .MajorUnitIsAuto = True
                .Crosses = xlMaximum
                .ReversePlotOrder = True
                .ScaleType = xlLinear
            End With
            With Selection
                .MajorTickMark = xlCross
                .MinorTickMark = xlOutside
                .TickLabelPosition = xlNextToAxis
            End With
            
            ActiveChart.PlotArea.Select
            With Selection.Border
                .Weight = xlThin
                .LineStyle = xlNone
            End With
            Selection.Interior.ColorIndex = xlNone
            
            ' ChartTitle Format
            ActiveChart.ChartTitle.Select
            GoSub chart_title_format
            
            ActiveChart.Axes(xlValue).AxisTitle.Select
            GoSub axis_title_format
            ActiveChart.Axes(xlCategory).AxisTitle.Select
            GoSub axis_title_format
            
            ActiveChart.Axes(xlCategory).Select
            GoSub axis_ticklabels_format
            ActiveChart.Axes(xlValue).Select
            GoSub axis_ticklabels_format
            
            ' chart page setup
            Call charts_sheet_pagesetup(ActiveChart)

            ' chart & plot area size
            Call charts_resize_chartarea
            
            ActiveWorkbook.Save
            ActiveWorkbook.Close
            
        Next g
    Next cp
    
    
    
    GoTo finish
    
chart_title_format:
        Selection.AutoScaleFont = False
        With Selection.Font
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
    Return
axis_title_format:
        Selection.AutoScaleFont = False
        With Selection.Font
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
    Return
axis_ticklabels_format:
        Selection.TickLabels.AutoScaleFont = False
        With Selection.TickLabels.Font
            .Name = "Times New Roman"
            .FontStyle = "Regular"
            .Size = 12
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
            .Background = xlAutomatic
        End With
    Return
    
finish:
    
End Sub
Sub ptsd_pet_plots_124overlay(plotPath, plotBook, plotSheet, dataPath, dataSheet, _
                               group, groupNames, groupFileNames, _
                               comp, _
                               beginTime, SampleRate, _
                               file_ext, _
                               title_ext, xaxis_title, yaxis_title)
'
' Macro coded 30/03/00 by Darren Weber
'

'
    
    For cp = 0 To UBound(comp)
        For g = 0 To UBound(group)
                        
            plotBook = (comp(cp) & file_ext & "_" & groupFileNames(g) & "_124overlay.xls")
    
            Workbooks.Add
            ActiveWorkbook.SaveAs FileName:=(plotPath & plotBook), FileFormat:=xlNormal
            Cells(1, 1).Select
            
            If (comp(cp) = "sa") Then
                cond = Array("ouc", "oac")
                condName = Array("Fixed Unattended Common", "Fixed Attended Common")
            ElseIf (comp(cp) = "wm") Then
                cond = Array("oac", "tac")
                condName = Array("Fixed Attended Common", "Variable Attended Common")
            ElseIf (comp(cp) = "ea") Then
                cond = Array("oac", "oat")
                condName = Array("Fixed Attended Common", "Fixed Attended Target")
            ElseIf (comp(cp) = "dt") Then
                cond = Array("tuc", "tud")
                condName = Array("Variable Unattended Common", "Variable Unattended Distractor")
            End If
            
            For C = 0 To UBound(condName)
                
                ' Open text data file and copy data
                dataSheet = (group(g) & cond(C) & file_ext & ".xls")
                Workbooks.Open FileName:=(dataPath & dataSheet)
                Range(Cells(1, 1), Cells(680, 124)).Select
                Selection.Copy
                
                ' Add new sheet and paste selection
                plotdataSheet = (group(g) & cond(C) & file_ext)
                Windows(plotBook).Activate
                Worksheets.Add.Name = plotdataSheet
                ActiveSheet.PASTE
                
                ' Save and close data file
                Workbooks(dataSheet).Save
                Workbooks(dataSheet).Close
                
                
                Worksheets(plotdataSheet).Select
                
                ' Copy last data row into next row
                Cells(681, 1).Activate
                ActiveCell.FormulaR1C1 = "=R[-1]C"
                Range(Cells(681, 1), Cells(681, 124)).Select
                Selection.FillRight
                
                ' Calculate column & overall min/max
                Cells(683, 1).Activate
                ActiveCell.FormulaR1C1 = "=MIN(R1C:R681C)"
                Cells(684, 1).Activate
                ActiveCell.FormulaR1C1 = "=MAX(R1C:R681C)"
                Range(Cells(683, 1), Cells(684, 124)).Select
                Selection.FillRight
                
                Cells(683, 126).Activate
                ActiveCell.FormulaR1C1 = "=MIN(RC1:RC124)"
                Cells(684, 126).Activate
                ActiveCell.FormulaR1C1 = "=MAX(RC1:RC124)"
                
                ' Add timescale data to column 126 (using formula and paste values)
                
                Cells(1, 126).Activate
                ActiveCell.FormulaR1C1 = beginTime
                Cells(2, 126).Activate
                ActiveCell.FormulaR1C1 = ("=R[-1]C + " & SampleRate)
                Cells(2, 126).Select
                Selection.AutoFill Destination:=Range(Cells(2, 126), Cells(681, 126)), Type:=xlFillDefault
                Range(Cells(2, 126), Cells(681, 126)).Select
                Selection.Copy
                Selection.PasteSpecial PASTE:=xlValues, Operation:=xlNone, SkipBlanks:= _
                    False, Transpose:=False
                Cells(1, 1).Select
                
                ' Add Chart
                Dim chartRange As Range
                
                Range(Cells(1, 1), Cells(681, 124)).Select
                Set chartRange = Range(Cells(1, 1), Cells(681, 124))
                Charts.Add
                ActiveChart.ChartType = xlLine
                ActiveChart.SetSourceData Source:=chartRange, _
                    PlotBy:=xlColumns
                
                For s = 1 To 124
                    ActiveChart.SeriesCollection(s).XValues = ("=" & plotdataSheet & "!" & "R1C126:R681C126")
                Next s
                
                ActiveChart.Location Where:=xlLocationAsNewSheet, Name:=(plotdataSheet & " plot")
                
                ' Format chart page setup
                Call charts_sheet_pagesetup(ActiveChart)
                
                ' Chart Title Definitions
                With ActiveChart
                    .HasTitle = True
                    .ChartTitle.Characters.Text = (groupNames(g) & ": " & condName(C) & title_ext)
                    .Axes(xlCategory, xlPrimary).HasTitle = True
                    .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = xaxis_title
                    .Axes(xlValue, xlPrimary).HasTitle = True
                    .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = yaxis_title
                End With
                
                ' Remove Legend
                ActiveChart.HasLegend = False
                
                ' Chart Axis Formats
                With ActiveChart
                    .HasAxis(xlCategory, xlPrimary) = True
                    .HasAxis(xlValue, xlPrimary) = True
                End With
                
                ActiveChart.Axes(xlCategory, xlPrimary).CategoryType = xlCategoryScale
                Call charts_add_gridlines
                
                ActiveChart.Axes(xlCategory).Select
                With Selection.Border
                    .Weight = xlHairline
                    .LineStyle = xlAutomatic
                End With
                With Selection
                    .MajorTickMark = xlCross
                    .MinorTickMark = xlOutside
                    .TickLabelPosition = xlNextToAxis
                End With
                With ActiveChart.Axes(xlCategory)
                    .crossesat = 1
                    .ticklabelspacing = 40
                    .tickmarkspacing = 40
                    .AxisBetweenCategories = False
                    .ReversePlotOrder = False
                End With
                ActiveChart.Axes(xlValue).Select
                With ActiveChart.Axes(xlValue)
                    .MinimumScaleIsAuto = True
                    .MaximumScaleIsAuto = True
                    .MinorUnitIsAuto = True
                    .MajorUnitIsAuto = True
                    .Crosses = xlMaximum
                    .ReversePlotOrder = True
                    .ScaleType = xlLinear
                End With
                With Selection
                    .MajorTickMark = xlCross
                    .MinorTickMark = xlOutside
                    .TickLabelPosition = xlNextToAxis
                End With
                
                ActiveChart.PlotArea.Select
                With Selection.Border
                    .Weight = xlThin
                    .LineStyle = xlNone
                End With
                Selection.Interior.ColorIndex = xlNone
                
                ' Chart Title Formats
                ActiveChart.ChartTitle.Select
                GoSub chart_title_format
                
                ActiveChart.Axes(xlValue).AxisTitle.Select
                GoSub axis_title_format
                ActiveChart.Axes(xlCategory).AxisTitle.Select
                GoSub axis_title_format
                
                ActiveChart.Axes(xlCategory).Select
                GoSub axis_ticklabels_format
                ActiveChart.Axes(xlValue).Select
                GoSub axis_ticklabels_format
                
                ' chart page setup
                Call charts_sheet_pagesetup(ActiveChart)

                ' chart & plot area size
                Call charts_resize_chartarea
            
            Next C
            
            ActiveWorkbook.Save
            ActiveWorkbook.Close
            
        Next g
    Next cp
    
    GoTo finish
    
chart_title_format:
        Selection.AutoScaleFont = True
        With Selection.Font
            .Name = "Times New Roman"
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
    Return
axis_title_format:
        Selection.AutoScaleFont = False
        With Selection.Font
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
    Return
axis_ticklabels_format:
        Selection.TickLabels.AutoScaleFont = False
        With Selection.TickLabels.Font
            .Name = "Times New Roman"
            .FontStyle = "Regular"
            .Size = 12
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
            .Background = xlAutomatic
        End With
    Return
    
finish:

End Sub
Sub ptsd_pet_plots_regions(plotPath, plotBook, plotSheet, dataPath, dataSheet, _
                           group, groupNames, groupFileNames, _
                           comp, _
                           beginTime, SampleRate, points, _
                           file_ext, _
                           title_ext, xaxis_title, yaxis_title)
'
' Macro recorded 30/03/00 by Darren Weber
'

'
    
    ' Define Regional Electrode Arrays
    
    Dim intI, intJ, region_elec(0 To 21, 0 To 5) As Integer, _
        regnames(21) As String
    
    region_names = Array("L_SPF", "L_IPF", "L_SF", "L_IF", "L_SC", "L_IC", "L_SP", "L_IP", "L_AT", "L_PT", "L_OT", _
                         "R_SPF", "R_IPF", "R_SF", "R_IF", "R_SC", "R_IC", "R_SP", "R_IP", "R_AT", "R_PT", "R_OT")
                         
    For intI = 0 To UBound(region_names)
        regnames(intI) = region_names(intI)
    Next intI
    
    L_SPF = Array(10, 11, 16, 17, 26, 0)
    For intJ = 0 To UBound(L_SPF)
            region_elec(0, intJ) = L_SPF(intJ)
    Next intJ
   
    L_IPF = Array(1, 4, 6, 7, 121, 122)
    For intJ = 0 To UBound(L_IPF)
            region_elec(1, intJ) = L_IPF(intJ)
    Next intJ
    
    L_SF = Array(25, 34, 35, 44, 45, 0)
    For intJ = 0 To UBound(L_SF)
            region_elec(2, intJ) = L_SF(intJ)
    Next intJ
    
    L_IF = Array(14, 15, 23, 24, 33, 0)
    For intJ = 0 To UBound(L_IF)
            region_elec(3, intJ) = L_IF(intJ)
    Next intJ
    
    L_SC = Array(53, 54, 62, 63, 72, 0)
    For intJ = 0 To UBound(L_SC)
            region_elec(4, intJ) = L_SC(intJ)
    Next intJ
    
    L_IC = Array(42, 43, 52, 60, 61, 0)
    For intJ = 0 To UBound(L_IC)
            region_elec(5, intJ) = L_IC(intJ)
    Next intJ
    
    L_SP = Array(71, 81, 82, 91, 99, 0)
    For intJ = 0 To UBound(L_SP)
            region_elec(6, intJ) = L_SP(intJ)
    Next intJ
    
    L_IP = Array(70, 79, 80, 89, 90, 98)
    For intJ = 0 To UBound(L_IP)
            region_elec(7, intJ) = L_IP(intJ)
    Next intJ
    
    L_AT = Array(31, 32, 50, 51, 0, 0)
    For intJ = 0 To UBound(L_AT)
            region_elec(8, intJ) = L_AT(intJ)
    Next intJ
    
    L_PT = Array(68, 69, 87, 88, 102, 111)
    For intJ = 0 To UBound(L_PT)
            region_elec(9, intJ) = L_PT(intJ)
    Next intJ
    
    L_OT = Array(103, 104, 109, 112, 116, 118)
    For intJ = 0 To UBound(L_OT)
            region_elec(10, intJ) = L_OT(intJ)
    Next intJ
    
    R_SPF = Array(12, 13, 19, 20, 27, 0)
    For intJ = 0 To UBound(R_SPF)
            region_elec(11, intJ) = R_SPF(intJ)
    Next intJ
    
    R_IPF = Array(3, 5, 8, 9, 123, 124)
    For intJ = 0 To UBound(R_IPF)
            region_elec(12, intJ) = R_IPF(intJ)
    Next intJ
    
    R_SF = Array(28, 37, 38, 46, 47, 0)
    For intJ = 0 To UBound(R_SF)
            region_elec(13, intJ) = R_SF(intJ)
    Next intJ
    
    R_IF = Array(21, 22, 29, 30, 39, 0)
    For intJ = 0 To UBound(R_IF)
            region_elec(14, intJ) = R_IF(intJ)
    Next intJ
    
    R_SC = Array(55, 56, 64, 65, 74, 0)
    For intJ = 0 To UBound(R_SC)
            region_elec(15, intJ) = R_SC(intJ)
    Next intJ
    
    R_IC = Array(48, 49, 57, 66, 67, 0)
    For intJ = 0 To UBound(R_IC)
            region_elec(16, intJ) = R_IC(intJ)
    Next intJ
    
    R_SP = Array(75, 83, 84, 93, 100, 0)
    For intJ = 0 To UBound(R_SP)
            region_elec(17, intJ) = R_SP(intJ)
    Next intJ
    
    R_IP = Array(76, 85, 86, 94, 95, 101)
    For intJ = 0 To UBound(R_IP)
            region_elec(18, intJ) = R_IP(intJ)
    Next intJ
    
    R_AT = Array(40, 41, 58, 59, 0, 0)
    For intJ = 0 To UBound(R_AT)
            region_elec(19, intJ) = R_AT(intJ)
    Next intJ
    
    R_PT = Array(77, 78, 96, 97, 108, 115)
    For intJ = 0 To UBound(R_PT)
            region_elec(20, intJ) = R_PT(intJ)
    Next intJ
    
    R_OT = Array(106, 107, 110, 114, 117, 120)
    For intJ = 0 To UBound(R_OT)
            region_elec(21, intJ) = R_OT(intJ)
    Next intJ
    
    L_SPF = "": L_IPF = "": L_SF = "":  L_IF = "": L_SC = ""
    L_IC = "":  L_SP = "":  L_IP = "":  L_AT = "": L_PT = ""
    L_OT = "":  R_SPF = "": R_IPF = "": R_SF = "": R_IF = ""
    R_SC = "":  R_IC = "":  R_SP = "":  R_IP = "": R_AT = ""
    R_PT = "":  R_OT = ""
    
    'For intI = 0 To UBound(region_elec)
    '    For intJ = 0 To 5
    '        reg_matrix = reg_matrix & " " & regnames(intI)
    '    Next intJ
    '        reg_matrix = reg_matrix & Chr(13)
    'Next intI
    'MsgBox (reg_matrix)
    
   
    ' Load data files
        
    For cp = 0 To UBound(comp)
        For g = 0 To UBound(group)
                        
            plotBook = (comp(cp) & file_ext & "_" & groupFileNames(g) & "_region_overlay.xls")
    
            Workbooks.Add
            ActiveWorkbook.SaveAs FileName:=(plotPath & plotBook), _
                FileFormat:=xlNormal, Password:="", WriteResPassword:="", _
                ReadOnlyRecommended:=False, CreateBackup:=False
            Cells(1, 1).Select
            
            If (comp(cp) = "sa") Then
                cond = Array("ouc", "oac")
                condName = Array("Fixed Unattended Common", "Fixed Attended Common")
            ElseIf (comp(cp) = "wm") Then
                cond = Array("oac", "tac")
                condName = Array("Fixed Attended Common", "Variable Attended Common")
            ElseIf (comp(cp) = "ea") Then
                cond = Array("oac", "oat")
                condName = Array("Fixed Attended Common", "Fixed Attended Target")
            ElseIf (comp(cp) = "dt") Then
                cond = Array("tuc", "tud")
                condName = Array("Variable Unattended Common", "Variable Unattended Distractor")
            End If
            
            
            For C = 0 To UBound(condName)
                
                ' Open text data file and copy data
                dataSheet = (group(g) & cond(C) & file_ext & ".xls")
                Workbooks.Open FileName:=(dataPath & dataSheet)
                Range(Cells(1, 1), Cells(680, 124)).Select
                Selection.Copy
                
                ' Add new sheet and paste selection
                difdataSheet = (group(g) & cond(C) & file_ext)
                Windows(plotBook).Activate
                Worksheets.Add.Name = difdataSheet
                ActiveSheet.PASTE
                
                ' Save and close data file
                Workbooks(dataSheet).Save
                Workbooks(dataSheet).Close
                
            Next C
            
            ' Add difference sheet and calculate difference values
            
            ASheet = (group(g) & cond(1) & file_ext)
            BSheet = (group(g) & cond(0) & file_ext)
            
            difdataSheet = (group(g) & cond(1) & "_vs_" & group(g) & cond(0) & file_ext)
            Worksheets.Add.Name = difdataSheet

            Cells(1, 1).Activate
            ActiveCell.FormulaR1C1 = ("=" & ASheet & "!RC - " & BSheet & "!RC")
            Range(Cells(1, 1), Cells(680, 1)).Select
            Selection.FillDown
            Range(Cells(1, 1), Cells(680, 124)).Select
            Selection.FillRight
            Selection.Copy
            Selection.PasteSpecial PASTE:=xlValues, Operation:=xlNone, SkipBlanks:= _
                False, Transpose:=False
            Cells(1, 1).Select
            
            ' Remove individual condition sheets (requires user input!)
            'Worksheets(ASheet).Delete
            'Worksheets(BSheet).Delete
            
            ' Copy last data row into next row
            Worksheets(difdataSheet).Select
            Cells(681, 1).Activate
            ActiveCell.FormulaR1C1 = "=R[-1]C"
            Range(Cells(681, 1), Cells(681, 124)).Select
            Selection.FillRight
            
            ' Calculate column & overall min/max
            Cells(683, 1).Activate
            ActiveCell.FormulaR1C1 = "=MIN(R1C:R681C)"
            Cells(684, 1).Activate
            ActiveCell.FormulaR1C1 = "=MAX(R1C:R681C)"
            Range(Cells(683, 1), Cells(684, 124)).Select
            Selection.FillRight
            
            Cells(683, 126).Activate
            ActiveCell.FormulaR1C1 = "=MIN(RC1:RC124)"
            Cells(684, 126).Activate
            ActiveCell.FormulaR1C1 = "=MAX(RC1:RC124)"
            
            ' Add timescale data to column 126 (using formula and paste values)
            Cells(1, 126).Activate
            ActiveCell.FormulaR1C1 = beginTime
            Cells(2, 126).Activate
            ActiveCell.FormulaR1C1 = ("=R[-1]C + " & SampleRate)
            Cells(2, 126).Select
            Selection.AutoFill Destination:=Range(Cells(2, 126), Cells(681, 126)), Type:=xlFillDefault
            Range(Cells(2, 126), Cells(681, 126)).Select
            Selection.Copy
            Selection.PasteSpecial PASTE:=xlValues, Operation:=xlNone, SkipBlanks:= _
                False, Transpose:=False
            Cells(1, 1).Select
                
            ' Add Charts
                    
            Dim chartRange As Range
                    
            For intI = 0 To UBound(region_elec)
                
                Col = region_elec(intI, 0)
                
                Worksheets(difdataSheet).Select
                Set chartRange = Range(Cells(1, Col), Cells(points, Col))

                Charts.Add
                ActiveChart.Name = regnames(intI)
                ActiveChart.ChartType = xlLine
                ActiveChart.SeriesCollection(1).Values = chartRange
                ActiveChart.SeriesCollection(1).XValues = ("=" & difdataSheet & "!" & "R1C126:R681C126")
                
                For intJ = 1 To 5
                    
                    Col = region_elec(intI, intJ)
                    
                    If (Col > 0) Then
                    
                        Worksheets(difdataSheet).Select
                        Set chartRange = Range(Cells(1, Col), Cells(points, Col))
                        
                        Charts(regnames(intI)).Activate
                        ActiveChart.SeriesCollection.NewSeries
                        ActiveChart.SeriesCollection(intJ + 1).Values = chartRange
                        ActiveChart.SeriesCollection(intJ + 1).XValues = ("=" & difdataSheet & "!" & "R1C126:R681C126")
                    End If
                
                Next intJ
                    
                ' Chart Title Definitions
                With ActiveChart
                    .HasTitle = True
                    .ChartTitle.Characters.Text = (groupNames(g) & ": " & comp(cp) & " dif: " & regnames(intI) & title_ext)
                    .Axes(xlCategory, xlPrimary).HasTitle = True
                    .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = xaxis_title
                    .Axes(xlValue, xlPrimary).HasTitle = True
                    .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = yaxis_title
                End With
                
                ' Remove Legend
                ActiveChart.HasLegend = False

                'Call chart_format
                
                'MsgBox "Excel is using: " & Application.MemoryUsed & " bytes" & Chr(13) & _
                       "Excel has free: " & Application.MemoryFree & " bytes"
                
            Next intI
            
            ActiveWorkbook.Save
            ActiveWorkbook.Close
            
        Next g
    Next cp

End Sub

Sub charts_format()

    Dim ch As Chart
    
    For Each ch In Charts
    
        ch.Activate

        ' Chart Axis Formats
        With ActiveChart
            .HasAxis(xlCategory, xlPrimary) = True
            .HasAxis(xlValue, xlPrimary) = True
        End With
        
        ActiveChart.Axes(xlCategory, xlPrimary).CategoryType = xlCategoryScale
        With ActiveChart.Axes(xlCategory)
            .HasMajorGridlines = False
            .HasMinorGridlines = False
        End With
        With ActiveChart.Axes(xlValue)
            .HasMajorGridlines = True
            .HasMinorGridlines = False
        End With
                        
        ActiveChart.Axes(xlCategory).Select
        With Selection.Border
            .Weight = xlHairline
            .LineStyle = xlAutomatic
        End With
        With Selection
            .MajorTickMark = xlCross
            .MinorTickMark = xlOutside
            .TickLabelPosition = xlNextToAxis
        End With
        With ActiveChart.Axes(xlCategory)
            .crossesat = 1
            .ticklabelspacing = 40
            .tickmarkspacing = 40
            .AxisBetweenCategories = False
            .ReversePlotOrder = False
        End With
        
        ActiveChart.Axes(xlValue).Select
        With ActiveChart.Axes(xlValue)
            .MinimumScaleIsAuto = True
            .MaximumScaleIsAuto = True
            .MinorUnitIsAuto = True
            .MajorUnitIsAuto = True
            .Crosses = xlMaximum    'xlAxisCrossesAutomatic
            .ReversePlotOrder = True
            .ScaleType = xlLinear
        End With
        With Selection
            .MajorTickMark = xlCross
            .MinorTickMark = xlOutside
            .TickLabelPosition = xlNextToAxis
        End With
        
        ActiveChart.PlotArea.Select
        With Selection.Border
            .Weight = xlHairline
            .LineStyle = xlNone
        End With
        With Selection.Interior
            .ColorIndex = 15
            .PatternColorIndex = 2
            .Pattern = xlSolid
        End With
        
        ' Chart Title Formats
        ActiveChart.ChartTitle.Select
        Selection.AutoScaleFont = True
        With Selection.Font
            .Name = "Arial"
            .Size = 20
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
            .Background = xlAutomatic
        End With
        
        ActiveChart.Axes(xlValue).AxisTitle.Select
        Selection.AutoScaleFont = True
        With Selection.Font
            .Name = "Arial"
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
        
        ActiveChart.Axes(xlCategory).AxisTitle.Select
        Selection.AutoScaleFont = True
        With Selection.Font
            .Name = "Arial"
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
        
        ActiveChart.Axes(xlCategory).Select
        Selection.TickLabels.AutoScaleFont = True
        With Selection.TickLabels.Font
            .Name = "Arial"
            .FontStyle = "Regular"
            .Size = 12
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
            .Background = xlAutomatic
        End With
        
        ActiveChart.Axes(xlValue).Select
        Selection.TickLabels.AutoScaleFont = True
        With Selection.TickLabels.Font
            .Name = "Arial"
            .FontStyle = "Regular"
            .Size = 12
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
            .Background = xlAutomatic
        End With

    Next ch

End Sub

Sub charts_remove_series3()
'
' remove_chart_series Macro
' Macro coded 27 Nov 2000 by Darren Weber
'
'
    
    Dim ch As ChartObject, n As Integer
    
    n = 0
    For Each ch In ActiveSheet.ChartObjects
    
        ch.Activate
        ch.Chart.SeriesCollection(3).Delete
        n = n + 1
    
    Next ch

    Dim Msg, Style, Title, Response
    Title = "Number of Charts"
    Msg = "Found " & n & " Charts."
    Style = vbOKOnly + vbInformation + vbDefaultButton3
    Response = MsgBox(Msg, Style, Title)

End Sub
Sub run_Yaxis_scale()

    'Call charts_edit_Yaxis_scale(-10, 10)
    'Return

    Dim Message, Title, min, max, major, minor, reverse
    
    Title = "Yaxis Scaling for All Charts"
    
    Message = "Do you want to rescale Yaxis of all chartobjects"
    Style = vbYesNoCancel + vbInformation + vbDefaultButton3
    Response = MsgBox(Message, Style, Title)
    
    If Response = vbYes Then
    
        Message = "Enter Yaxis minimum"
        min = InputBox(Message, Title, "-10")
        
        Message = "Enter Yaxis maximum"
        max = InputBox(Message, Title, "10")
        
        Message = "Enter Yaxis majorunits"
        major = InputBox(Message, Title, "1")
        
        Message = "Enter Yaxis minorunits"
        minor = InputBox(Message, Title, "0.5")
        
        Message = "Reverse Yaxis Polarity"
        reverse = InputBox(Message, Title, "True")
    
        Call charts_edit_Yaxis_scale(min, max, major, minor, reverse)
    End If
    
End Sub
Sub run_Xaxis_scale()

    'Call charts_edit_Xaxis_scale(40)
    'Return

    Dim Message, Title
    
    Title = "Xaxis Scaling for All Charts"
    
    Message = "Do you want to rescale Xaxis of all chartobjects"
    Style = vbYesNoCancel + vbInformation + vbDefaultButton3
    Response = MsgBox(Message, Style, Title)
    
    If Response = vbYes Then
    
        Message = "Enter Xaxis tickmark spacing"
        tickmarkspacing = InputBox(Message, Title, "10")
        
        Message = "Yaxis crosses at category number?"
        crossesat = InputBox(Message, Title, "10")
        
        Message = "Enter Xaxis ticklabel spacing"
        ticklabelspace = InputBox(Message, Title, "20")
        
        'Message = "Reverse Yaxis Polarity"
        'reverse = InputBox(Message, Title, "True")
        
        Call charts_edit_Xaxis_scale(tickmarkspacing, crossesat, ticklabelspace)

    End If
End Sub

Sub charts_edit_Xaxis_scale(tickmarkspacing, _
                            Optional crossesat = 80, _
                            Optional ticklabelspacing = 80, _
                            Optional Yaxisbetweencat = False, _
                            Optional reverse = False)

    n = 0
    
    Dim ch As ChartObject
    For Each ch In ActiveSheet.ChartObjects
    
    'Dim ch As Chart, n As Integer
    'For Each ch In Charts
    
        
        'Create/Disply category xaxis
        ch.Chart.HasAxis(xlCategory, xlPrimary) = True
        
        ch.Chart.Axes(xlCategory, xlPrimary).CategoryType = xlCategoryScale
        With ch.Chart.Axes(xlCategory)
            
            .crossesat = crossesat
            .ticklabelspacing = ticklabelspacing
            .tickmarkspacing = tickmarkspacing
            .AxisBetweenCategories = Yaxisbetweencat
            .ReversePlotOrder = reverse
            
        End With
                
        n = n + 1
        
    Next ch
    
    Call charts_processed_report(n)
    
End Sub


Sub charts_edit_Xaxis_properties()
'
' charts_edit_Xaxis_properties Macro
' Macro coded 27 Mar 2001 by Darren Weber
'
'
    n = 0
    
    Dim ch As ChartObject
    For Each ch In ActiveSheet.ChartObjects
    
    'Dim ch As Chart, n As Integer
    'For Each ch In Charts
    
        
        'Create/Disply category xaxis
        ch.Chart.HasAxis(xlCategory, xlPrimary) = True
        
        ch.Chart.Axes(xlCategory, xlPrimary).CategoryType = xlCategoryScale
        With ch.Chart.Axes(xlCategory)
            
            .HasMajorGridlines = False
            .HasMinorGridlines = False

            .Border.Weight = xlHairline
            .Border.LineStyle = xlAutomatic 'xlNone
        
            .MajorTickMark = xlCross
            .MinorTickMark = xlNone
            .TickLabelPosition = xlNone 'xlNextToAxis
            
            .TickLabels.AutoScaleFont = False
            With .TickLabels.Font
                .Name = "Times New Roman"
                .FontStyle = "Regular"
                .Size = 10
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

    Call charts_processed_report(n)

End Sub

Sub charts_edit_Yaxis_scale(min, max, _
                            Optional majorunit = 1, _
                            Optional minorunit = 0.5, _
                            Optional reverse = True)
'
' charts_edit_Yaxis_scale Macro
' Macro coded 27 Nov 2000 by Darren Weber
'
'
    n = 0
    
    Dim ch As ChartObject
    For Each ch In ActiveSheet.ChartObjects
    
    'Dim ch As Chart, n As Integer      ' For chart sheets
    'For Each ch In Charts
    
        ch.Activate
        
        With ActiveChart.Axes(xlValue)
            .MinimumScale = min
            .MaximumScale = max
            .majorunit = majorunit
            .minorunit = minorunit
            .ReversePlotOrder = reverse
            
            .ScaleType = xlLinear
            .Crosses = xlAxisCrossesAutomatic 'xlMaximum
        
        End With
        
        n = n + 1
    
    Next ch

    Call charts_processed_report(n)

End Sub

Sub charts_edit_Yaxis_properties()
'
' charts_edit_Yaxis_properties Macro
' Macro coded 27 Nov 2000 by Darren Weber
'
'
    
    Dim ch As ChartObject, n As Integer
    
    n = 0
    For Each ch In ActiveSheet.ChartObjects
    
        With ch.Chart.Axes(xlValue)
        
            .Border.Weight = xlHairline
            .Border.LineStyle = xlNone 'xlAutomatic ' xlNone for no axis
            
            .MajorTickMark = xlNone 'xlCross
            .MinorTickMark = xlNone 'xlOutside
            .TickLabelPosition = xlNone 'xlNextToAxis
        
            .TickLabels.AutoScaleFont = False
            With .TickLabels.Font
                .Name = "Times New Roman"
                .FontStyle = "Regular"
                .Size = 10
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

    Call charts_processed_report(n)

End Sub

Sub charts_resize_and_format(Optional h = 80, Optional w = 95)
'
' charts_resize_and_format Macro
' Macro coded 27 Nov 2000 by Darren Weber
'       08/2001 added chart axis scaling &
'               top/left indents, titles handling, & other formats
'
' Warning: reduce runtime by minimising excel window first!
'

    Lindent = 0
    
    goty = False
    gotx = False

    Dim ch As ChartObject, n As Integer
    n = 0
    For Each ch In ActiveSheet.ChartObjects
        ch.Activate
        With ch
        
            If .Chart.HasTitle = True Then
                .Name = "Chart " & .Chart.ChartTitle.Characters.Text
                With .Chart.ChartTitle
                    .Top = 0
                    Tindent = .Font.Size + 5
                    
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
            

            ' Get chart x-axis & y-axis properties and create scale
            If n = 0 Then
                Call charts_resize_scale(1000, 150, ch, 10, 5, " ms", "uV")
            End If

        End With
        n = n + 1
    Next ch
    
    Call charts_processed_report(n)

End Sub

Sub charts_edit_series_properties()
'
' charts_edit_series_properties Macro
' Macro coded 27/11/00 by Darren Weber
'
    
    Dim ch As ChartObject
    
    For Each ch In ActiveSheet.ChartObjects
        ch.Activate
        With ch
        
            For Series = 1 To 2
        
                
                'Define Series Properties
                '------------------------
                Select Case Series
                    Case 1
                        sColor = 3
                        sWeight = xlHairline    ' xlHairline, xlThin, xlMedium, or xlThick
                        sLine = xlContinuous
                        sPlotOrder = 1
                    Case 2
                        sColor = 5
                        sWeight = xlHairline
                        sLine = xlContinuous
                End Select
                
                'Set Series Properties
                '---------------------
                With .Chart.SeriesCollection(Series).Border
                    .ColorIndex = sColor
                    .Weight = sWeight
                    .LineStyle = sLine
                End With
                With .Chart.SeriesCollection(Series)
                    
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

Sub charts_edit_series_legend_properties()

    xsize = 75
    xorigin = 175
    yorigin = 150

    SeriesLines = Array("Line Series1", "Line Series2")
    'Remove current series lines
    '--------------------------
    For Each sName In SeriesLines
        On Error Resume Next
        ActiveSheet.Shapes(sName).Delete
    Next sName
    
    'Create New Horizontal Series Lines
    '----------------------------------
    With ActiveSheet.Shapes.AddLine(xorigin, yorigin, xorigin + xsize, yorigin)
        .Name = SeriesLines(0)
        .Line.Weight = 5#
        .Line.DashStyle = msoLineSolid
        .Line.Style = msoLineSingle
        .Line.ForeColor.SchemeColor = 10 ' Red
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
    
    With ActiveSheet.Shapes.AddLine(xorigin, yorigin + 25, xorigin + xsize, yorigin + 25)
        .Name = SeriesLines(1)
        .Line.Weight = 5#
        .Line.DashStyle = msoLineSolid
        .Line.Style = msoLineSingle
        .Line.ForeColor.SchemeColor = 48 ' Light Blue
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

Sub charts_edit_series_rows_range()
'
' edit_series_range_rows Macro
' Macro coded 9/04/01 by Darren Weber
'

'

    Dim Msg, Style, Title, Response
    Title = "Chart series formula"
    Style = vbOKOnly + vbInformation + vbDefaultButton3

    Dim ch As ChartObject, _
        formula, old_start_row, new_start_row, old_end_row, new_end_row  As String, _
        new_srow, new_erow As Integer
    
    old_start_row = "$1:"
    new_start_row = "$1:"
    new_xstart_row = 0      'set to 0 to skip xrange update
    old_end_row = "$425,"
    new_end_row = "$401,"
    new_xend_row = 401
    
    
    For Each ch In ActiveSheet.ChartObjects
        
        ch.Activate
        With ch
            
            formula = .Chart.SeriesCollection(1).formula
            GoSub REPLACE_SERIES_ROWS
            .Chart.SeriesCollection(1).formula = formula
            
            If new_xstart_row > 0 Then
                Xrange = "R" & new_xstart_row & "C126:" & "R" & new_xend_row & "C126"
                .Chart.SeriesCollection(1).XValues = ("='" & ActiveWorkbook.Sheets(1) & "'!" & Xrange)
            End If
            
            'MsgBox formula
            
            formula = .Chart.SeriesCollection(2).formula
            GoSub REPLACE_SERIES_ROWS
            .Chart.SeriesCollection(2).formula = formula
            
            'MsgBox formula
            
        End With
        
    Next ch
   
    GoTo Fin
    
REPLACE_SERIES_ROWS:

    X = 1
    While (InStr(formula, old_start_row) > 0) And (X < 5)
        Mid(formula, InStr(formula, old_start_row)) = new_start_row
        X = X + 1
    Wend
        
    X = 1
    While (InStr(formula, old_end_row) > 0) And (X < 5)
        Mid(formula, InStr(formula, old_end_row)) = new_end_row
        X = X + 1
            
        Msg = "Formula is: " & formula
        'Response = MsgBox(Msg, Style, Title)
            
    Wend
    Return
    
Fin:
End Sub


Sub chart_lineplot_add2sheet()
'
' chart_lineplot_add2sheet Macro
' Macro coded 19/01/01 by Darren Weber
'

'
    Application.CutCopyMode = False
    Charts.Add
    ActiveChart.ChartType = xlLine
    ActiveChart.SetSourceData Source:=Sheets("Sheet1").Range("A1:B680"), PlotBy _
        :=xlColumns
    ActiveChart.Location Where:=xlLocationAsObject, Name:="Sheet2"
    With ActiveChart
        .HasTitle = False
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "msec"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "uV/m2"
    End With
    With ActiveChart
        .HasAxis(xlCategory, xlPrimary) = True
        .HasAxis(xlValue, xlPrimary) = True
    End With
    ActiveChart.Axes(xlCategory, xlPrimary).CategoryType = xlAutomatic
    With ActiveChart.Axes(xlCategory)
        .HasMajorGridlines = False
        .HasMinorGridlines = False
    End With
    With ActiveChart.Axes(xlValue)
        .HasMajorGridlines = False
        .HasMinorGridlines = False
    End With
    ActiveChart.HasLegend = False
    ActiveChart.HasDataTable = False
    ActiveSheet.Shapes("Chart 1").IncrementLeft -97.8
    ActiveSheet.Shapes("Chart 1").IncrementTop -57.6
End Sub

Sub charts_resize_title()
'
' charts_resize_title Macro
' Macro coded 08/01 by Darren Weber
'
'
    Dim ch As ChartObject, n As Integer
    n = 0
    For Each ch In ActiveSheet.ChartObjects
        ch.Activate
        If ch.Chart.HasTitle = True Then
            
            With ch
            
                With .Chart.ChartTitle
                    .HorizontalAlignment = xlHAlignCenter 'xlLeft
                    .VerticalAlignment = xlVAlignCenter 'xlTop
                    .Orientation = xlHorizontal
                End With
                .Chart.ChartTitle.AutoScaleFont = False
                With .Chart.ChartTitle.Font
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
                chTitleSize = .Chart.ChartTitle.Font.Size
                chMiddle = (ch.Width / 2) - ((chTitleWidth / 2) * chTitleSize)
                
                .Chart.ChartTitle.Top = 1
                .Chart.ChartTitle.left = chMiddle
                'Top right
                'shwidth = .Chart.ChartTitle.Width
                '.Chart.ChartTitle.left = chwidth - (shwidth + 5)
                
            End With
            n = n + 1
        End If
    Next ch

    Dim Msg, Style, Title, Response
    Title = "Number of Charts"
    Msg = "Modified Size of " & n & " Charts."
    Style = vbOKOnly + vbInformation + vbDefaultButton3
    'Response = MsgBox(Msg, Style, Title)

End Sub

Sub charts_resize_scale(xorigin, yorigin, ch, _
                        xk, yk, _
                        Optional x_scale_text = " ms", _
                        Optional y_scale_text = "uV")
'
' charts_size_scale Macro
' Macro coded 15/08/01 by Darren Weber
'
'

    With ch
        If .Chart.HasAxis(xlCategory) = True Then
            With .Chart.Axes(xlCategory)
                
                cNames = .CategoryNames
                        
                cfirst = LBound(cNames)
                clast = UBound(cNames)
                nCategories = clast - cfirst
                        
                ' requires definition of msec range in 'special' series
                x_min = cNames(cfirst)
                x_max = cNames(clast)
                        
                x_sample_rate = cNames(cfirst + 1) - x_min
                x_spacing = .tickmarkspacing
                x_units = x_sample_rate * x_spacing
                x_length = .Width
                x_pixels = x_length / ((x_max - x_min) / x_units)
                        
            End With
        End If
        If .Chart.HasAxis(xlValue) = True Then
            With .Chart.Axes(xlValue)
                y_units = .majorunit
                y_min = .MinimumScale
                y_max = .MaximumScale
                y_length = .Height
                y_pixels = y_length / ((y_max - y_min) / y_units)
                goty = True
            End With
        End If
    End With
    
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
    
    ' Add x scale line
    '----------------
    With ActiveSheet.Shapes.AddLine(xorigin, yorigin, xorigin + x_pixels, yorigin)
        .Name = "x_scale_line"
        ' set other line properties here
    End With
    
    'Add x label
    '-----------
    x_label_left = xorigin + 20
    x_label_top = yorigin + 5
    x_scale_text = CStr(x_min) & " - " & CStr(x_max) & x_scale_text
    
    With ActiveSheet.Shapes.AddLabel(msoTextOrientationHorizontal, x_label_left, x_label_top, 0#, 0#)
        .Name = "x_scale_label"
        .TextFrame.Characters.Text = x_scale_text
    End With
        
    'Add y scale line
    '----------------
    baseline = x_min * (x_pixels / x_units)
    If x_min < 0 Then
        yx_origin = xorigin - baseline
    Else
        yx_origin = xorigin
    End If
    y_start = yorigin + y_pixels
    y_end = yorigin - y_pixels
    
    With ActiveSheet.Shapes.AddLine(yx_origin, y_start, yx_origin, y_end)
        .Name = "y_scale_line"
    End With
    
    'Add amplitude labels
    '--------------------
    y_label_left = xorigin - 50
    y_label_top = y_end
    
    With ActiveSheet.Shapes.AddLabel(msoTextOrientationHorizontal, y_label_left, y_label_top, 0#, 0#)
        .TextFrame.Characters.Text = y_scale_text
        .Name = "y_scale_label"
    End With
    
    py_label_left = xorigin
    py_label_top = y_start
    py_scale_text = "+" & CStr(y_units)
    With ActiveSheet.Shapes.AddLabel(msoTextOrientationHorizontal, py_label_left, py_label_top, 0#, 0#)
        .TextFrame.Characters.Text = py_scale_text
        .Name = "py_scale_label"
    End With
    
    ny_label_left = xorigin
    ny_label_top = y_end - 20
    ny_scale_text = "-" & CStr(y_units)
    With ActiveSheet.Shapes.AddLabel(msoTextOrientationHorizontal, ny_label_left, ny_label_top, 0#, 0#)
        .TextFrame.Characters.Text = ny_scale_text
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
            With .Characters.Font
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

Sub charts_add_gridlines()

    With ActiveChart.Axes(xlCategory)
        .HasMajorGridlines = True
        .HasMinorGridlines = False
        With .MajorGridlines.Border
            .ColorIndex = 15        ' Gray 25%
            .Weight = xlHairline
            .LineStyle = xlContinuous
        End With
    End With
    
    With ActiveChart.Axes(xlValue)
        .HasMajorGridlines = True
        .HasMinorGridlines = False
        With .MajorGridlines.Border
            .ColorIndex = 15        ' Gray 25%
            .Weight = xlHairline
            .LineStyle = xlContinuous
        End With
    End With

End Sub

Sub charts_resize_chartarea(Optional w = 30, Optional h = 21)
    
    ' w & h are in cm and set to A4 page size by default
    
    If ActiveChart.PageSetup.ChartSize = xlFullPage Then
        ' Can't do anything
    Else
    
        Lindent = 0
        
        With ActiveChart
        
            .ChartArea.Width = .Application.CentimetersToPoints(w)
            .ChartArea.Height = .Application.CentimetersToPoints(h)
            
            'If .HasTitle = True Then
            '    With .ChartTitle
            '        .Top = 0
            '        TopIndent = (.Font.Size * 2) + 5
            '    End With
            'Else
            '    TopIndent = 0
            'End If
            
            'Hplot = .ChartArea.Height
            'Wplot = .ChartArea.Width
            
            '.PlotArea.Top = TopIndent
            '.PlotArea.left = Lindent
            '.PlotArea.Height = Hplot
            '.PlotArea.Width = Wplot
            '
            'If .HasAxis(xlValue) = True Then
            '    If .Axes(xlValue).HasTitle = True Then
            '        Linsideplot = .Axes(xlValue).AxisTitle.Font.Size + 5
            '    End If
            'End If
            'If .HasAxis(xlCategory) = True Then
            '    If .Axes(xlCategory).HasTitle = True Then
            '        Hinsideplot = .Axes(xlCategory).AxisTitle.Font.Size + 5
            '    End If
            'End If
            
        End With
    End If
End Sub

