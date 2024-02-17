Attribute VB_Name = "SetDateInTableColumn"
'@Folder "Code Common"
Option Explicit
Public Sub SetDateInTableColumn(ByVal sourceTable As ListObject, ByRef columnTitle As String)

    Dim sourceColumn As ListColumn
    Set sourceColumn = sourceTable.ListColumns.[_Default](columnTitle)
    Dim DateField As Object
    If Not sourceColumn Is Nothing Then
        For Each DateField In sourceColumn.DataBodyRange
            If IsError(DateField) Then DateField = Now
        Next
    End If
    
End Sub
