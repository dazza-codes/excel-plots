VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SeriesRange 
   Caption         =   "Series Range"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3210
   OleObjectBlob   =   "SeriesRange.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SeriesRange"
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    
    SeriesRange.LabelOldStart.Caption = ERP_Plot70.series_oldstart
    SeriesRange.LabelOldEnd.Caption = ERP_Plot70.series_oldend
    SeriesRange.TextBoxNewStart.Value = ERP_Plot70.series_start
    SeriesRange.TextBoxNewEnd.Value = ERP_Plot70.series_end
End Sub
    
Private Sub CommandButtonCancel_Click()
    ERP_Plot70.bRun = False
    Unload Me
End Sub

Private Sub CommandButtonOK_Click()

    ERP_Plot70.series_start = SeriesRange.TextBoxNewStart.Value
    ERP_Plot70.series_end = SeriesRange.TextBoxNewEnd.Value
    
    ERP_Plot70.bRun = True
    Unload Me
End Sub

