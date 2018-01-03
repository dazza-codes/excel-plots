VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SeriesProperties 
   Caption         =   "Series Properties"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3375
   OleObjectBlob   =   "SeriesProperties.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SeriesProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub ComboBoxS1Label_Change()
    
    selectSeries = SeriesProperties.ComboBoxS1Label.Value
    
    Dim ch As ChartObject, s As Integer
    
    For Each ch In ActiveSheet.ChartObjects
        
        For Each Series In ch.Chart.SeriesCollection
            
            'Set Series Properties
            '---------------------
            If InStr(1, Series.Name, selectSeries, vbTextCompare) > 0 Then
                
                'Get Series Properties
                '---------------------
                With ch.Chart.SeriesCollection(Series.Name).Border
                    sColor = .Color
                    sWeight = .Weight
                    sLine = .LineStyle
                    
                End With
                
                sPlotOrder = ch.Chart.SeriesCollection(Series.Name).PlotOrder
                
                GoTo Fin
                
            End If
            
        Next Series
        
    Next ch
    
Fin:
    
    ' Update other combo boxes
    SeriesProperties.TextBoxS1NewLabel.Value = selectSeries
    SeriesProperties.TextBoxS1Position.Value = sPlotOrder
    SeriesProperties.TextBoxS1Weight.Value = sWeight
    
    Select Case sLine
        Case xlLineStyleNone
            SeriesProperties.ComboBoxS1Line.Value = "None"
        Case xlAutomatic
            SeriesProperties.ComboBoxS1Line.Value = "Continuous"
        Case xlContinuous
            SeriesProperties.ComboBoxS1Line.Value = "Continuous"
        Case xlDash
            SeriesProperties.ComboBoxS1Line.Value = "Dash"
        Case xlDashDot
            SeriesProperties.ComboBoxS1Line.Value = "DashDot"
        Case xlDashDotDot
            SeriesProperties.ComboBoxS1Line.Value = "DashDotDot"
        Case xlDot
            SeriesProperties.ComboBoxS1Line.Value = "Dot"
        Case xlDouble
            SeriesProperties.ComboBoxS1Line.Value = "Double"
        Case xlSlantDashDot
            SeriesProperties.ComboBoxS1Line.Value = "SlantDashDot"
    End Select
    
    'Select Case sWeight
    '    Case xlHairline
    '        SeriesProperties.ComboBoxS1Weight.Value = "Hairline"
    '    Case xlThin
    '        SeriesProperties.ComboBoxS1Weight.Value = "Thin"
    '    Case xlMedium
    '        SeriesProperties.ComboBoxS1Weight.Value = "Medium"
    '    Case xlThick
    '        SeriesProperties.ComboBoxS1Weight.Value = "Thick"
    'End Select
    
    Select Case sColor
        Case vbBlack
            SeriesProperties.ComboBoxS1Color.Value = "Black"
        Case vbRed
            SeriesProperties.ComboBoxS1Color.Value = "Red"
        Case vbGreen
            SeriesProperties.ComboBoxS1Color.Value = "Green"
        Case vbYellow
            SeriesProperties.ComboBoxS1Color.Value = "Yellow"
        Case vbBlue
            SeriesProperties.ComboBoxS1Color.Value = "Blue"
        Case vbMagenta
            SeriesProperties.ComboBoxS1Color.Value = "Magenta"
        Case vbCyan
            SeriesProperties.ComboBoxS1Color.Value = "Cyan"
        Case vbWhite
            SeriesProperties.ComboBoxS1Color.Value = "White"
        Case RGB(191, 191, 191)
            SeriesProperties.ComboBoxS1Color.Value = "25% Grey"
        Case RGB(128, 128, 128)
            SeriesProperties.ComboBoxS1Color.Value = "50% Grey"
        Case RGB(65, 65, 65)
            SeriesProperties.ComboBoxS1Color.Value = "75% Grey"
    End Select
    

End Sub

Private Sub UserForm_Initialize()
    
    
    SeriesProperties.ComboBoxS1Label.ColumnCount = 1
    SeriesProperties.ComboBoxS1Label.ColumnWidths = 15
    
    Dim ch As ChartObject, s As Integer
    For Each ch In ActiveSheet.ChartObjects
        
        For Each Series In ch.Chart.SeriesCollection
            
            'Set Series Properties
            '---------------------
            With ch.Chart.SeriesCollection(Series.Name).Border
                sColor = .Color
                sWeight = .Weight
                sLine = .LineStyle
            End With
            
            SeriesProperties.ComboBoxS1Label.AddItem Series.Name
            
        Next Series
        
        GoTo Init
        
    Next ch
    
Init:
    
    SeriesProperties.ComboBoxS1Line.ColumnCount = 1
    SeriesProperties.ComboBoxS1Line.ColumnWidths = 15
    SeriesProperties.ComboBoxS1Line.AddItem "None"
    SeriesProperties.ComboBoxS1Line.AddItem "Continuous"
    SeriesProperties.ComboBoxS1Line.AddItem "Dash"
    SeriesProperties.ComboBoxS1Line.AddItem "DashDot"
    SeriesProperties.ComboBoxS1Line.AddItem "DashDotDot"
    SeriesProperties.ComboBoxS1Line.AddItem "Dot"
    SeriesProperties.ComboBoxS1Line.AddItem "Double"
    SeriesProperties.ComboBoxS1Line.AddItem "SlantDashDot"
    
    SeriesProperties.ComboBoxS1Color.ColumnCount = 1
    SeriesProperties.ComboBoxS1Color.AddItem "White"
    SeriesProperties.ComboBoxS1Color.AddItem "25% Grey"
    SeriesProperties.ComboBoxS1Color.AddItem "50% Grey"
    SeriesProperties.ComboBoxS1Color.AddItem "75% Grey"
    SeriesProperties.ComboBoxS1Color.AddItem "Black"
    SeriesProperties.ComboBoxS1Color.AddItem "Red"
    SeriesProperties.ComboBoxS1Color.AddItem "Green"
    SeriesProperties.ComboBoxS1Color.AddItem "Yellow"
    SeriesProperties.ComboBoxS1Color.AddItem "Blue"
    SeriesProperties.ComboBoxS1Color.AddItem "Magenta"
    SeriesProperties.ComboBoxS1Color.AddItem "Cyan"
    
End Sub

Private Sub CommandOK_Click()
    
    ERP_Plot70.s1Label = SeriesProperties.ComboBoxS1Label.Value
    ERP_Plot70.s1NewLabel = SeriesProperties.TextBoxS1NewLabel.Value
    ERP_Plot70.s1PlotOrder = SeriesProperties.TextBoxS1Position.Value
    
    linetext = SeriesProperties.ComboBoxS1Line.Value
    Select Case linetext
        Case "None"
            linevalue = xlLineStyleNone
        Case "Automatic"
            linevalue = xlAutomatic
        Case "Continuous"
            linevalue = xlContinuous
        Case "Dash"
            linevalue = xlDash
        Case "DashDot"
            linevalue = xlDashDot
        Case "DashDotDot"
            linevalue = xlDashDotDot
        Case "Dot"
            linevalue = xlDot
        Case "Double"
            linevalue = xlDouble
        Case "SlantDashDot"
            linevalue = xlSlantDashDot
    End Select
    ERP_Plot70.s1Line = linevalue
    
    
    ERP_Plot70.s1Weight = SeriesProperties.TextBoxS1Weight.Value
    
    'Select Case SeriesProperties.ComboBoxS1Weight.Value
    '    Case "Hairline"
    '        ERP_Plot70.s1Weight = xlHairline
    '    Case "Thin"
    '        ERP_Plot70.s1Weight = xlThin
    '    Case "Medium"
    '        ERP_Plot70.s1Weight = xlMedium
    '    Case "Thick"
    '        ERP_Plot70.s1Weight = xlThick
    'End Select
    
    coltext = SeriesProperties.ComboBoxS1Color.Value
    Select Case coltext
        Case "Black"
            colval = RGB(0, 0, 0)       'vbBlack
        Case "Red"
            colval = RGB(255, 0, 0)     'vbRed
        Case "Green"
            colval = RGB(0, 255, 0)     'vbGreen
        Case "Yellow"
            colval = RGB(255, 255, 0)   'vbYellow
        Case "Blue"
            colval = RGB(0, 0, 255)     'vbBlue
        Case "Magenta"
            colval = RGB(255, 0, 255)   'vbMagenta
        Case "Cyan"
            colval = RGB(0, 255, 255)   'vbCyan
        Case "White"
            colval = RGB(255, 255, 255) 'vbWhite
        Case "25% Grey"
            colval = RGB(191, 191, 191)
        Case "50% Grey"
            colval = RGB(128, 128, 128)
        Case "75% Grey"
            colval = RGB(65, 65, 65)
    End Select
    ERP_Plot70.s1Color = colval
    
    ERP_Plot70.bRun = True
    Unload Me
End Sub

Private Sub CommandCancel_Click()
    ERP_Plot70.bRun = False
    Unload Me
End Sub
