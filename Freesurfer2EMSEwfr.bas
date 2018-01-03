Attribute VB_Name = "Freesurfer"
Sub load_freesurfer_tri()
Attribute load_freesurfer_tri.VB_Description = "Macro recorded 4/19/00 by Darren Weber"
Attribute load_freesurfer_tri.VB_ProcData.VB_Invoke_Func = " \n14"
'
' load_freesurfer_tri Macro
' Macro coded 4/19/00 by Darren Weber
'
'
    '##############################################################
    '
    ' This macro converts a freesurfer triangulation file
    ' (eg, brain.tri in bem area of anatomical analysis)
    ' into an EMSE wireframe file (.wfr)
    '
    '##############################################################
    
    ' Manually Import text file using delimited space parsing and save as .xls file
    
    
    ' Declare parameters
    Dim vertices, triangles, Row As Integer, col2Temp As Double, FName As String
    
    FName = ActiveWorkbook.FullName
    
    ' Remove the first column, if empty
    If (StrComp(Cells(1, 1), "") = 0) Then
        Columns(1).Select
        Selection.Delete Shift:=xlToLeft
    End If
    
    ' Define number of vertices and delete row
    vertices = Cells(1, 1)
    Rows(1).Select
    Selection.Delete Shift:=xlUp
    
    ' Add file header information. For details, see:
    ' http://www.sourcesignal.com/fileform/Wireframe.htm
    
    Rows(1).Select
    Selection.Insert Shift:=xlDown
    Selection.Insert Shift:=xlDown
    Selection.Insert Shift:=xlDown
    
    'EMSE format header information
    Cells(1, 1) = 3
    Cells(1, 2) = 4000
    Cells(2, 1) = 3      ' signifies the minor revision for this file format
    
    ' 40:scalp, 80:outer skull, 100:inner skull, 200:cortex
    If ((InStr(FName, "skin") > 0)) Then
        Cells(3, 1) = 40
    ElseIf ((InStr(FName, "outer_skull") > 0)) Then
        Cells(3, 1) = 80
    ElseIf ((InStr(FName, "inner_skull") > 0)) Then
        Cells(3, 1) = 100
    ElseIf ((InStr(FName, "cortex") > 0)) Then
        Cells(3, 1) = 200
    Else
        Cells(3, 1) = 40 ' default to scalp
    End If
   
    
    ' For each vertex: _

    '    (a) replace vertex number with "v" character _

    '    (b) swap col 2 value with col 3 value _

    '    EMSE's convention uses a right hand system, _
            i.e. if you point your four fingers together in the direction _
            of the sequence 1,2,3 of the corners of a triangle, your thumb _
            will point outwards.
    
    Row = 4
    For v = 1 To vertices
        Cells(Row, 1) = "v"
        col2Temp = Cells(Row, 2)
        Cells(Row, 2) = Cells(Row, 3)
        Cells(Row, 3) = col2Temp
        Row = Row + 1
    Next v
        
    ' Define number of triangles and delete row
    triangles = Cells(Row, 1)
    Rows(Row).Select
    Selection.Delete Shift:=xlUp
    
    ' For each triangle, replace triangle number with "t" character
    For tri = 1 To triangles
        Cells(Row, 1) = "t"
        Cells(Row, 2) = Cells(Row, 2) - 1
        Cells(Row, 3) = Cells(Row, 3) - 1
        Cells(Row, 4) = Cells(Row, 4) - 1
        Row = Row + 1
    Next tri
    
    Cells(1, 1).Select
    
    
    ' Substitute .xls for .wfr file extension
    Start = InStr(FName, ".tri")
    Mid(FName, Start) = ".wfr"
    
    MsgBox "Saving the active workbook into EMSE wireframe file " & FName
    
    ActiveWorkbook.SaveAs FileName:=FName, FileFormat:=xlText, _
        CreateBackup:=False
    
End Sub
