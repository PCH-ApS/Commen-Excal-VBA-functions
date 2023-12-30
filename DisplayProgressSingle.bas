Public Sub DisplayProgressSingle(pctCompl1 As Long, pctText1 As String, timeText3 As String)

CommonSingleProgress.Text.Caption = pctText1 & ": " & Int(pctCompl1) & "%"
CommonSingleProgress.Bar.Width = pctCompl1 * 2
CommonSingleProgress.Text3.Caption = "Elapsed time: " & timeText3
DoEvents

End Sub