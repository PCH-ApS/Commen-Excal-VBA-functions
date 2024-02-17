Attribute VB_Name = "SelectSingleFile"
Option Explicit
'@Folder "Code Common"
Public Function SelectSingleFile(ByVal dialogTitle As String, ByVal fileFilter As String, ByVal fileType As String) As String
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .title = dialogTitle
        .Filters.Clear
        .Filters.Add fileFilter, fileType
        If .Show = True Then
                SelectSingleFile = .SelectedItems.Item(1)
            Else
                Exit Function
        End If
    End With

End Function
