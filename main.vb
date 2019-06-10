Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Function Main()
    Dim imagePath As String
    imagePath = BrowseFile
    If Len(imagePath) > 0 Then
        Call PrintImage(imagePath, 1)
    End If
End Function

Function PrintImage(ByVal imagePath As String, _
                    ByVal imageScale As Double)
    Dim wia As Object
    Dim p As StdPicture
    Dim y, x, w, h As Integer
    Dim r, g, b As Integer
    Dim c As Long
       
    Call ClearAll
    
    Set wia = CreateObject("WIA.ImageFile")
    wia.LoadFile imagePath
    w = wia.Width
    h = wia.Height
    
    Call SetCellSize(w, h, imageScale)
    
    Set p = LoadPicture(imagePath)
    hdc = CreateCompatibleDC(0)
    SelectObject hdc, p.Handle

    For y = 1 To h Step 1
        For x = 1 To w Step 1
            c = GetPixel(hdc, x - 1, y - 1)
                Cells(y, x).Interior.Color = c
        Next x
    Next y
End Function

Function BrowseFile() As String
    
    Dim fP As String 'filePath
    Dim fd As Office.fileDialog
    
     Set fd = Application.fileDialog(msoFileDialogFilePicker)
    
     With fd
        .AllowMultiSelect = False
        .Title = "Select image"
        .Filters.Clear
        .Filters.Add "Bitmap", "*.bmp"
        
        If .Show = True Then
            fP = Dir(.SelectedItems(1))
        End If
     End With
     
     BrowseFile = fP
End Function

's = imageScale
Function SetCellSize(ByVal w As Integer, _
                    ByVal h As Integer, _
                    ByVal s As Double)
                    
                    
    With Range(Cells(1, 1), Cells(h, w))
        .ColumnWidth = 0.1 * s
        .RowHeight = 1 * s
    End With
End Function

Function ClearAll()
    Cells.Select
    Selection.Clear
    Selection.ColumnWidth = 8.43
    Selection.RowHeight = 15
    Cells(1, 1).Select
End Function
