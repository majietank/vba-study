'**********************************************************************************************
'       制作总表模块
'**********************************************************************************************
'按钮
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
'----------------------------------------------------------------------------------------------
'调用
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
'----------------------------------------------------------------------------------------------
'调用
Sub 车牌行()
Dim i As Integer
Dim chepai As String
For i = 1 To Sheets.Count - 1
    chepai = Sheets(i).Name
    Cells(2, i + 1) = chepai
Next i
Rows(2).HorizontalAlignment = xlCenter
Rows(2).VerticalAlignment = xlCenter
Columns(1).HorizontalAlignment = xlCenter
Columns(1).VerticalAlignment = xlCenter
End Sub
'----------------------------------------------------------------------------------------------
'调用
Sub 记录信息(riqi As Date)
Cells(3, "O") = Year(riqi) & "年"
Cells(3, "Q") = Month(riqi) & "月"
Cells(47, "E") = Null
Cells(47, "M") = Null
Cells(47, "P") = Null
End Sub
'**********************************************************************************************
'       单元格格式模块
'**********************************************************************************************
'按钮
Sub 表格格式初始化()
Dim i As Integer
For i = 1 To Sheets.Count - 1
With Sheets(i).Activate
    Cells(1, 1).Select
    'ActiveWindow.View = xlPageBreakPreview
    ActiveWindow.View = xlPageLayoutView
    Columns("B:S").ColumnWidth = 3.35
    Columns("A").ColumnWidth = 8.38
    Rows("1:1").RowHeight = 25
    Rows("2:51").RowHeight = 15
    ActiveWindow.Zoom = 100
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
    Range("A1:S51").Font.Name = "宋体"
    Range("A1:S51").NumberFormatLocal = "G/通用格式"
    Range("A1:S51").Cells.Replace What:=" ", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Rows("48:51").RowHeight = 15
    Cells(48, "A").Orientation = xlVertical
    Cells(5, "A").Orientation = xlVertical
    LeftMargin = Application.InchesToPoints(0.708661417322835)
    RightMargin = Application.InchesToPoints(0.708661417322835)
    TopMargin = Application.InchesToPoints(0.354330708661417)
    BottomMargin = Application.InchesToPoints(0.275590551181102)
    HeaderMargin = Application.InchesToPoints(0.31496062992126)
    FooterMargin = Application.InchesToPoints(0.31496062992126)
End With
Next i
End Sub
'**********************************************************************************************
'       打印、存储模块
'**********************************************************************************************
'按钮
Sub 打印存档()
Dim i As Integer
Dim owjm As String, wjm As String
wjm = "运行记录" & Sheet99.Cells(1, 2) & "打印存储.xlsx"
owjm = ThisWorkbook.Name
Workbooks.Add
ActiveWorkbook.SaveAs Filename:=Workbooks(owjm).Path & "\" & wjm, FileFormat:=xlWorkbookDefault, CreateBackup:=False
Workbooks(owjm).Activate
For i = 1 To Sheets.Count - 1
    If Not Sheets(i).Range("S16:S46").Find("*保") Is Nothing Then
        Sheets(i).Copy before:=Workbooks(wjm).Sheets("Sheet1")
        Workbooks(owjm).Activate
    End If
Next i
Workbooks(wjm).Activate
Sheets("Sheet1").Delete
ActiveWorkbook.Save
ActiveWorkbook.Close
End Sub
'----------------------------------------------------------------------------------------------
'按钮
Sub 打印表格()
Dim i As Integer
For i = 1 To Sheets.Count - 1
Sheets(i).Activate
If Not Range("S16:S46").Find("*保") Is Nothing Then
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
End If
Sheet99.Activate
Next i
End Sub
'----------------------------------------------------------------------------------------------
'按钮
Sub 记录存档()
Dim i As Integer
Dim chepai As String
Sheet99.Select
Dim wjm As String, owjm As String
owjm = ThisWorkbook.Name
wjm = "运行记录" & Sheet99.Cells(1, 2) & ".xlsx"
Workbooks.Add
ActiveWorkbook.SaveAs Filename:=Workbooks(owjm).Path & "\" & wjm, FileFormat:=xlWorkbookDefault, CreateBackup:=False
Workbooks(owjm).Activate
For i = 1 To Sheets.Count - 1
    chepai = Sheets(i).Name
    Sheets(chepai).Select
    Workbooks(owjm).Sheets(chepai).Copy before:=Workbooks(wjm).Sheets("Sheet1")
    Windows(owjm).Activate
Next i
Workbooks(wjm).Activate
Sheets("Sheet1").Delete
ActiveWorkbook.Save
ActiveWorkbook.Close
End Sub
'**********************************************************************************************
'                                   记录录入模块
'**********************************************************************************************
'按钮
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
'----------------------------------------------------------------------------------------------
'调用
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
'                                          新增设备模块
'**********************************************************************************************
'按钮
Sub 新增设备()
Dim chepai As String
chepai = InputBox("输入车牌：")
Sheets(1).Copy before:=Sheet99
With ActiveSheet
    .Name = chepai
End With
Sheet99.Activate
Call 车牌行
End Sub