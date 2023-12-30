Public Function SelectSaveDest(ByVal filename As String, ByVal fileFilter As String, ByVal dialogTitle As String) As String
    
    SelectSaveDest = Application.GetSaveAsFilename(fileFilter:=fileFilter, Title:=dialogTitle, InitialFileName:=filename)
    
End Function