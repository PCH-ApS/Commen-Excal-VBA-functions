Public Sub DisplayProgressOverall(pctCompl1 As Long, pctCompl2 As Long, pctText1 As String, pctText2 As String, timeText3 As String)

CommonOverallProgress.Text.Caption = pctText1 & ": " & Int(pctCompl1) & "%"
CommonOverallProgress.Bar.Width = pctCompl1 * 2
CommonOverallProgress.Text2.Caption = pctText2 & ": " & Int(pctCompl2) & "%"
CommonOverallProgress.Bar2.Width = pctCompl2 * 2
CommonOverallProgress.Text3.Caption = "Elapsed time: " & timeText3
DoEvents

End Sub