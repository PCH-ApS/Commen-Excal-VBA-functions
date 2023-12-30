Public Function SelectMultiFile(ByVal dialogTitle As String, ByVal fileFilter As String, ByVal fileType As String) As Variant

    Dim fncSelectedItem() As Variant
    Dim filesSelected As Long
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = True
        .Title = dialogTitle
        .Filters.Clear
        .Filters.Add fileFilter, fileType
        If .Show = True Then
                ReDim fncSelectedItem(1 To .SelectedItems.Count)
                For filesSelected = 1 To .SelectedItems.Count
                    fncSelectedItem(filesSelected) = .SelectedItems(filesSelected)
                Next
                SelectMultiFile = fncSelectedItem()
            Else
                Exit Function
        End If
    End With

End Function