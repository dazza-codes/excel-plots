VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SeriesAdd 
   Caption         =   "Series Add"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   OleObjectBlob   =   "SeriesAdd.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SeriesAdd"
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    
    SeriesAdd.ComboBoxSheetsAvailable.ColumnCount = 1
    
    For Each sheet In ActiveWorkbook.Worksheets
        
        SeriesAdd.ComboBoxSheetsAvailable.AddItem sheet.Name
        
    Next sheet
    
End Sub

Private Sub CommandOK_Click()
    
    ERP_Plot70.series_sheet = SeriesAdd.ComboBoxSheetsAvailable.Value
    
    ERP_Plot70.bRun = True
    Unload Me
End Sub

Private Sub CommandCancel_Click()
    ERP_Plot70.bRun = False
    Unload Me
End Sub

