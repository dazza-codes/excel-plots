VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChartsScale 
   Caption         =   "Charts Scale"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3420
   OleObjectBlob   =   "ChartsScale.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ChartsScale"
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()

    Dim ch As ChartObject
    For Each ch In ActiveSheet.ChartObjects
        With ch
            ChartsScale.ChartHeight.Value = CLng(.Height)
            ChartsScale.ChartWidth.Value = CLng(.Width)
        End With
        GoTo Init
    Next ch
    
Init:
    
    ' Get Scale Location etc., if present
    ScaleShapes = Array("x_scale_line", "x_scale_metric", "y_scale_metric")
    For Each shName In ScaleShapes
        On Error Resume Next
        Select Case shName
            Case ScaleShapes(0)
            ChartsScale.XOrigin = CLng(ActiveSheet.Shapes(shName).Left)
            ChartsScale.YOrigin = CLng(ActiveSheet.Shapes(shName).Top)
            Case ScaleShapes(1)
                ChartsScale.ScaleXLabel = ActiveSheet.Shapes(shName).TextFrame.Characters.Text
            Case ScaleShapes(2)
                ChartsScale.ScaleYLabel = ActiveSheet.Shapes(shName).TextFrame.Characters.Text
        End Select
    Next shName
    
    ERP_Plot70.bRun = False
    
End Sub

Private Sub CommandOK_Click()
    
    ERP_Plot70.lChartHeight = CLng(ChartsScale.ChartHeight.Value)
    ERP_Plot70.lChartWidth = CLng(ChartsScale.ChartWidth.Value)
    ERP_Plot70.lScaleXOrigin = CLng(ChartsScale.XOrigin.Value)
    ERP_Plot70.lScaleYOrigin = CLng(ChartsScale.YOrigin.Value)
    ERP_Plot70.sScaleXLabel = ChartsScale.ScaleXLabel.Value
    ERP_Plot70.sScaleYLabel = ChartsScale.ScaleYLabel.Value
    
    ERP_Plot70.bRun = True
    Unload Me
    
End Sub

Private Sub CommandCancel_Click()
    
    ERP_Plot70.bRun = False
    Unload Me
    
End Sub
