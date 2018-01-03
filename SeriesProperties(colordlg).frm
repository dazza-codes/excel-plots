VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SeriesProperties 
   Caption         =   "Series Properties"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4665
   OleObjectBlob   =   "SeriesProperties(colordlg).frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SeriesProperties"
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    
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
    
            'Define Series Properties
            '------------------------
            Select Case Series.Name
                Case "Series1"
                    ERP_Plot70.s1Color = sColor
                    ERP_Plot70.s1Weight = sWeight
                    ERP_Plot70.s1Line = sLine
                Case "Series2"
                    ERP_Plot70.s2Color = sColor
                    ERP_Plot70.s2Weight = sWeight
                    ERP_Plot70.s2Line = sLine
            End Select
            s = s + 1
        Next Series
        
        GoTo Init
    
    Next ch
    
Init:
    
    SeriesProperties.ComboBoxS1Line.ColumnCount = 1
    SeriesProperties.ComboBoxS1Line.ColumnWidths = 15
    SeriesProperties.ComboBoxS1Line.AddItem "None"
    SeriesProperties.ComboBoxS1Line.AddItem "Automatic"
    SeriesProperties.ComboBoxS1Line.AddItem "Continuous"
    SeriesProperties.ComboBoxS1Line.AddItem "Dash"
    SeriesProperties.ComboBoxS1Line.AddItem "DashDot"
    SeriesProperties.ComboBoxS1Line.AddItem "DashDotDot"
    SeriesProperties.ComboBoxS1Line.AddItem "Dot"
    SeriesProperties.ComboBoxS1Line.AddItem "Double"
    SeriesProperties.ComboBoxS1Line.AddItem "SlantDashDot"
    
    SeriesProperties.ComboBoxS2Line.ColumnCount = 1
    SeriesProperties.ComboBoxS2Line.ColumnWidths = 15
    SeriesProperties.ComboBoxS2Line.AddItem "None"
    SeriesProperties.ComboBoxS2Line.AddItem "Automatic"
    SeriesProperties.ComboBoxS2Line.AddItem "Continuous"
    SeriesProperties.ComboBoxS2Line.AddItem "Dash"
    SeriesProperties.ComboBoxS2Line.AddItem "DashDot"
    SeriesProperties.ComboBoxS2Line.AddItem "DashDotDot"
    SeriesProperties.ComboBoxS2Line.AddItem "Dot"
    SeriesProperties.ComboBoxS2Line.AddItem "Double"
    SeriesProperties.ComboBoxS2Line.AddItem "SlantDashDot"
    
    For Line = 1 To 2
    
        If Line = 1 Then
            linevalue = ERP_Plot70.s1Line
        Else
            linevalue = ERP_Plot70.s2Line
        End If
        
        Select Case linevalue
            Case xlLineStyleNone
                linetext = "None"
            Case xlAutomatic
                linetext = "Automatic"
            Case xlContinuous
                linetext = "Continuous"
            Case xlDash
                linetext = "Dash"
            Case xlDashDot
                linetext = "DashDot"
            Case xlDashDotDot
                linetext = "DashDotDot"
            Case xlDot
                linetext = "Dot"
            Case xlDouble
                linetext = "Double"
            Case xlSlantDashDot
                linetext = "SlantDashDot"
        End Select
        
        If Line = 1 Then
            SeriesProperties.ComboBoxS1Line.Value = linetext
        Else
            SeriesProperties.ComboBoxS2Line.Value = linetext
        End If
    Next Line
    
    SeriesProperties.ComboBoxS1Weight.ColumnCount = 1
    SeriesProperties.ComboBoxS1Weight.AddItem "Hairline"
    SeriesProperties.ComboBoxS1Weight.AddItem "Thin"
    SeriesProperties.ComboBoxS1Weight.AddItem "Medium"
    SeriesProperties.ComboBoxS1Weight.AddItem "Thick"
    
    SeriesProperties.ComboBoxS2Weight.ColumnCount = 1
    SeriesProperties.ComboBoxS2Weight.AddItem "Hairline"
    SeriesProperties.ComboBoxS2Weight.AddItem "Thin"
    SeriesProperties.ComboBoxS2Weight.AddItem "Medium"
    SeriesProperties.ComboBoxS2Weight.AddItem "Thick"
    
    For Weight = 1 To 2
    
        If Weight = 1 Then
            weightvalue = ERP_Plot70.s1Weight
        Else
            weightvalue = ERP_Plot70.s2Weight
        End If
        
        Select Case weightvalue
            Case xlHairline
                weighttext = "Hairline"
            Case xlThin
                weighttext = "Thin"
            Case xlMedium
                weighttext = "Medium"
            Case xlThick
                weighttext = "Thick"
        End Select
        
        If Weight = 1 Then
            SeriesProperties.ComboBoxS1Weight.Value = weighttext
        Else
            SeriesProperties.ComboBoxS2Weight.Value = weighttext
        End If
    Next Weight
    
    SeriesProperties.CommonDialogS1Color.Flags = cdlCCRGBInit
    SeriesProperties.CommonDialogS1Color.Color = s1Color
    SeriesProperties.CommandButtonS1Color.BackColor = SeriesProperties.CommonDialogS1Color.Color
    SeriesProperties.CommonDialogS2Color.Flags = cdlCCRGBInit
    SeriesProperties.CommonDialogS2Color.Color = s2Color
    SeriesProperties.CommandButtonS2Color.BackColor = SeriesProperties.CommonDialogS2Color.Color
    
End Sub

Private Sub CommandButtonS1Color_Click()
    SeriesProperties.CommonDialogS1Color.ShowColor
    SeriesProperties.CommandButtonS1Color.BackColor = SeriesProperties.CommonDialogS1Color.Color
End Sub

Private Sub CommandButtonS2Color_Click()
    SeriesProperties.CommonDialogS2Color.ShowColor
    SeriesProperties.CommandButtonS2Color.BackColor = SeriesProperties.CommonDialogS2Color.Color
End Sub

Private Sub CommandOK_Click()
    
    For Line = 1 To 2
    
        If Line = 1 Then
            linetext = SeriesProperties.ComboBoxS1Line.Value
        Else
            linetext = SeriesProperties.ComboBoxS2Line.Value
        End If
      
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
    
        If Line = 1 Then
            ERP_Plot70.s1Line = linevalue
        Else
            ERP_Plot70.s2Line = linevalue
        End If
    
    Next Line
    
    Select Case SeriesProperties.ComboBoxS1Weight.Value
        Case "Hairline"
            ERP_Plot70.s1Weight = xlHairline
        Case "Thin"
            ERP_Plot70.s1Weight = xlThin
        Case "Medium"
            ERP_Plot70.s1Weight = xlMedium
        Case "Thick"
            ERP_Plot70.s1Weight = xlThick
    End Select
    
    Select Case SeriesProperties.ComboBoxS2Weight.Value
        Case "Hairline"
            ERP_Plot70.s2Weight = xlHairline
        Case "Thin"
            ERP_Plot70.s2Weight = xlThin
        Case "Medium"
            ERP_Plot70.s2Weight = xlMedium
        Case "Thick"
            ERP_Plot70.s2Weight = xlThick
    End Select
    
    ERP_Plot70.s1Color = SeriesProperties.CommonDialogS1Color.Color
    ERP_Plot70.s2Color = SeriesProperties.CommonDialogS2Color.Color
    
    ERP_Plot70.bRun = True
    Unload Me
End Sub

Private Sub CommandCancel_Click()
    ERP_Plot70.bRun = False
    Unload Me
End Sub
