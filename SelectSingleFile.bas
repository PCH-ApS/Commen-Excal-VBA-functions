Public Function SelectSingleFile(ByVal dialogTitle As String, ByVal fileFilter As String, ByVal fileType As String) As String
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Title = dialogTitle
        .Filters.Clear
        .Filters.Add fileFilter, fileType
        If .Show = True Then
                SelectSingleFile = .SelectedItems(1)
            Else
                Exit Function
        End If
    End With

End Function