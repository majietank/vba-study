Sub 表格数据初始化()
Dim i As Integer
Dim chepai As String
Dim riqi As Date
Sheet99.Activate
'判断日期
If Day(Date) > 25 Then
    riqi = Date
Else
    riqi = Date - Day(Date)
End If
Sheet99.Cells(1, 1) = Year(riqi) & "年"
Sheet99.Cells(1, 2) = Month(riqi) & "月"
Call 车牌行
Call 日期列(riqi, 3)
'循环表
For i = 1 To Sheets.Count - 1
    With Sheets(i).Activate
        Call 记录信息(riqi)
        Call 日期列(riqi, 16)
        Range("B16:S46") = Null
    End With
    Sheet99.Activate
Next i
End Sub
Sub 日期列(riqi As Date, qishi As Integer)
Dim i As Integer
Dim riqirow As Date
For i = 1 To 31
    If i <= 25 Then
        riqirow = DateSerial(Year(riqi), Month(riqi), i)
        Cells(i + qishi + 5, 1) = Day(riqirow)
    Else
        riqirow = DateSerial(Year(riqi), Month(riqi) - 1, i)
        If Month(riqi) = Month(riqirow) Then
            Cells(i - 26 + qishi, 1) = Null
        Else
            Cells(i - 26 + qishi, 1) = Day(riqirow)
        End If
    End If
Next i
End Sub
Sub 车牌行()
Dim i As Integer
Dim chepai As String
For i = 1 To Sheets.Count - 1
    chepai = Sheets(i).Name
    Cells(2, i + 1) = chepai
Next i
End Sub
Sub 记录信息(riqi As Date)
Cells(3, "O") = Year(riqi) & "年"
Cells(3, "Q") = Month(riqi) & "月"
Cells(47, "E") = Null
Cells(47, "M") = Null
Cells(47, "P") = Null
End Sub
'**********************************************************************************************
Sub 表格格式初始化()
Dim i As Integer
For i = 1 To Sheets.Count - 1
With Sheets(i).Activate
    Cells(1, 1).Select
    ActiveWindow.View = xlPageBreakPreview
    Columns("B:S").ColumnWidth = 3.38
    Columns("A").ColumnWidth = 8.38
    ActiveWindow.Zoom = 100
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
    Range("A1:S51").Font.Name = "宋体"
    Range("A1:S51").NumberFormatLocal = "G/通用格式"
    Range("A1:S51").Cells.Replace What:=" ", Replacement:="", LookAt:=xlPart, SearchOrder:= _
        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Rows("48:51").RowHeight = 15
    Cells(48, "A").Orientation = xlVertical
    Cells(5, "A").Orientation = xlVertical
End With
Next i
End Sub
'**********************************************************************************************
Sub 打印表格()
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
End Sub
Sub 记录存档()
Dim i As Integer
Dim chepai As String
Sheet99.Select
Dim wjm As String, owjm As String
owjm = ThisWorkbook.Name
wjm = "运行记录" & Sheet99.Cells(1, 2) & ".xlsx"
Workbooks.Add
ActiveWorkbook.SaveAs Filename:=wjm, FileFormat:=xlWorkbookDefault, CreateBackup:=False
Windows(owjm).Activate
For i = 1 To Sheets.Count - 1
    chepai = Sheets(i).Name
    Sheets(chepai).Select
    Sheets(chepai).Copy Before:=Workbooks(wjm).Sheets(1)
Next i
With Windows(wjm).Activate
Sheets("Sheet1").Delete
ActiveWorkbook.Save
ActiveWindow.Close
End With
End Sub
'**********************************************************************************************
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
'**********************************************************************************************
Sub 新增设备()
Dim chepai As String
chepai = InputBox("输入车牌：")
Sheets(1).Copy Before:=Sheet99
With ActiveSheet
    .Name = chepai
End With
Sheet99.Activate
Call 车牌行
End Sub



