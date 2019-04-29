Sub 记录录入()
Dim i As Integer, chewei As Integer
Dim chepai As String
Sheet99.Activate
For i = 1 To Sheets.Count - 1
    chepai = Sheets(i).Name
    chewei = Sheet99.Rows(2).Find(chepai).Column
    Range(Cells(3, chewei), Cells(33, chewei)).Copy
    With Sheets(i).Activate
        Range("P16").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Call 记录判断
    End With
    Sheet99.Activate
Next i
End Sub
Sub 记录判断()
Dim i  As Integer
For i = 16 To 46
If Cells(i, 1) <> 0 Then
    If Cells(i, "P") <> 0 Then
        Range(Cells(i, "B"), Cells(i, "O")) = "√"
        Cells(i, "S") = Null
    Else
        Cells(i, "S") = "○"
    End If
End If
Next i
End Sub
