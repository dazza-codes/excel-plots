VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} XaxisScale 
   Caption         =   "Category(X) Axis Scales"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3945
   OleObjectBlob   =   "XaxisScale.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "XaxisScale"
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()
    
    Dim ch As ChartObject
    For Each ch In ActiveSheet.ChartObjects
        
        If ch.Chart.HasAxis(xlCategory) Then
            With ch.Chart.Axes(xlCategory)
                ERP_Plot70.lTickMarkSp = .TickMarkSpacing
                ERP_Plot70.lTickLabelSp = .TickLabelSpacing
                ERP_Plot70.dCrossesAt = .CrossesAt
                ERP_Plot70.bYBetween = .AxisBetweenCategories
                ERP_Plot70.bXreverse = .ReversePlotOrder
                
            End With
        End If
        GoTo Init
    Next ch
    
Init:
    XaxisScale.TickMarkSpacing.Value = ERP_Plot70.lTickMarkSp
    XaxisScale.TickLabelSpacing.Value = ERP_Plot70.lTickLabelSp
    XaxisScale.CrossesAt.Value = ERP_Plot70.dCrossesAt
    XaxisScale.CheckBoxCrossBetween.Value = ERP_Plot70.bYBetween
    XaxisScale.CheckBoxXReverse.Value = ERP_Plot70.bXreverse
    
End Sub

Private Sub CommandOK_Click()
    
    ERP_Plot70.lTickMarkSp = XaxisScale.TickMarkSpacing.Value
    ERP_Plot70.lTickLabelSp = XaxisScale.TickLabelSpacing.Value
    
    ERP_Plot70.dCrossesAt = XaxisScale.CrossesAt.Value
    
    ERP_Plot70.bYBetween = XaxisScale.CheckBoxCrossBetween.Value
    ERP_Plot70.bXreverse = XaxisScale.CheckBoxXReverse.Value
    
    ERP_Plot70.bRun = True
    Unload Me
End Sub

Private Sub CommandCancel_Click()
    
    ERP_Plot70.bRun = False
    Unload Me
End Sub

