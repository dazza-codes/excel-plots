VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} YaxisScale 
   Caption         =   "Value(Y) Axis Scales"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2175
   OleObjectBlob   =   "YaxisScale.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "YaxisScale"
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()

    Dim ch As ChartObject
    For Each ch In ActiveSheet.ChartObjects
        
        With ch.Chart.Axes(xlValue)
            ERP_Plot70.dMin = .MinimumScale
            ERP_Plot70.dMax = .MaximumScale
            ERP_Plot70.dMajor = .MajorUnit
            ERP_Plot70.dMinor = .MinorUnit
            ERP_Plot70.bYreverse = .ReversePlotOrder
        End With
        GoTo Init
    Next ch
    
Init:
    YaxisScale.Maximum.Value = ERP_Plot70.dMax
    YaxisScale.Minimum.Value = ERP_Plot70.dMin
    YaxisScale.MajorUnit.Value = ERP_Plot70.dMajor
    YaxisScale.MinorUnit.Value = ERP_Plot70.dMinor
    YaxisScale.ReverseValues.Value = ERP_Plot70.bYreverse
    
End Sub

Private Sub CommandOK_Click()
    
    ERP_Plot70.dMax = YaxisScale.Maximum.Value
    ERP_Plot70.dMin = YaxisScale.Minimum.Value
    ERP_Plot70.dMajor = YaxisScale.MajorUnit.Value
    ERP_Plot70.dMinor = YaxisScale.MinorUnit.Value
    ERP_Plot70.bYreverse = YaxisScale.ReverseValues.Value
    
    ERP_Plot70.bRun = True
    Unload Me
    
End Sub

Private Sub CommandCancel_Click()
    
    ERP_Plot70.bRun = False
    Unload Me
    
End Sub
