'**********************************************************************************************
'       总表模块
'**********************************************************************************************
'////////////记录子模块\\\\\\\\\\\\\\
'总表按钮
Option Explicit
Public riqi As Date
Sub 总表生成()
Dim i As Integer
Dim chepai As String
Dim jilu As Range
Sheet99.Activate
    '设置总表第一行
Rows("1:1").RowHeight = 25
    '设置所有单元格居中
With Cells
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With
    '判断日期
If Day(Date) > 25 Then
    riqi = Date
Else
    riqi = Date - Day(Date)
End If
    '填写总表日期
Sheet99.Cells(1, 1) = Year(riqi) & "年"
Sheet99.Cells(1, 2) = Month(riqi) & "月"
    '车牌行制作
Call 车牌行
    '日期列制作
Call 日期列(riqi, 3)
Set jilu = Range(Cells(3, 2), Cells(33, Sheets.Count - 1))
jilu.ClearContents
Rows("35:35").RowHeight = 25
End Sub
'----------------------------------------------------------------------------------------------
'全局调用
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
'总表调用
Sub 车牌行()
Dim i As Integer
Dim chepai As String
For i = 1 To Sheets.Count - 1
    chepai = Sheets(i).Name
    Cells(2, i + 1) = chepai
Next i
End Sub
'////////////油耗子模块\\\\\\\\\\\\\\
'按钮
Sub 油耗表生成()
Dim i As Integer, j As Integer
Windows("当月油耗统计表.xlsx").Activate
Columns(1).ColumnWidth = 20
j = Range(Cells(36, 1), Cells(65536, 1)).End(xlDown).Row
For i = 36 To j
If Cells(i, 3) <> 0 Then
    Cells(i, 5) = Cells(i, 3) / Cells(i, 4)
End If
Next i
End Sub
Sub 油耗列()
Dim i As Integer, chewei As Integer
Dim youhao As Range
Dim chepai As String
Sheet99.Activate
For i = 36 To 72
'车牌列
Set youhao = Range(Cells(i, 1), Cells(i, 2))
youhao.Select
With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
End With

'**********************************************************************************************
'       记录表模块
'**********************************************************************************************
'----------------------------------------------------------------------------------------------
Sub 循环选择记录表()
Dim i As Integer
For i = 1 To Sheets.Count - 1
    With Sheets(i).Activate
        Call 记录表信息(riqi)
        Call 日期列(riqi, 16)
        Range("B16:S46") = Null
    End With
    Sheet99.Activate
Next i
End Sub
'调用
Sub 记录表信息(riqi As Date)
Cells(3, "O") = Year(riqi) & "年"
Cells(3, "Q") = Month(riqi) & "月"
Cells(47, "E") = Null
Cells(47, "M") = Null
Cells(47, "P") = Null
End Sub

'按钮
Sub 台班统计()
Dim i As Integer
Dim chejilu As Range
For i = 1 To Sheets.Count - 1
    Set chejilu = Range(Cells(3, i + 1), Cells(33, i + 1))
    Cells(34, i + 1) = WorksheetFunction.CountA(chejilu)
Next i
End Sub
'**********************************************************************************************
'       计算每日油耗
'**********************************************************************************************

'**********************************************************************************************
'       单元格格式模块
'**********************************************************************************************
'按钮
Sub 表格格式初始化()
Dim i As Integer
For i = 1 To Sheets.Count - 1
With Sheets(i).Activate
'    ActiveWindow.View = xlPageBreakPreview
    ActiveWindow.View = xlPageLayoutView
    ActiveWindow.Zoom = 100
    Columns("A:A").ColumnWidth = 8.38
    Columns("B:S").ColumnWidth = 3.38
    Rows("1:1").RowHeight = 25
    Rows("2:51").RowHeight = 14.5
'    ActiveSheet.PageSetup.LeftMargin = Application.InchesToPoints(0.7)
'    ActiveSheet.PageSetup.RightMargin = Application.InchesToPoints(0.7)
'    ActiveSheet.PageSetup.TopMargin = Application.InchesToPoints(0.5)
'    ActiveSheet.PageSetup.BottomMargin = Application.InchesToPoints(0.5)
    Range("A1:S51").Font.Name = "新宋体"
    Range("A1:S1").Font.Size = 20
    Range("A2:S47").Font.Size = 11
    Range("A48").Font.Size = 11
    Range("C48:S51").Font.Size = 9
    Cells(1, 1).Select
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

'**********************************************************************************************
Sub test()
Dim i%, b%
Dim riqirange0 As Range
Dim riqirange1 As Range
Dim riqirange2 As Range
Dim riqirange3 As Range
'清理子表所有内容
Cells.Delete
Set riqirange0 = Range(Cells(3, 1), Cells(33, 1))
Call 日期列(3)
'循环子表
For i = 1 To Sheets.Count - 2
Cells(2, i * 3) = Sheets(i).Name
Range(Cells(2, i * 3), Cells(2, i * 3 + 2)).Select
With Selection
    .Merge
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Interior.ThemeColor = xlThemeColorAccent3
    .Interior.TintAndShade = 0.2
End With
Range(Cells(34, i * 3), Cells(34, i * 3 + 2)).Select
With Selection
    .Merge
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Interior.ThemeColor = xlThemeColorAccent3
    .Interior.TintAndShade = 0.2
End With
Columns(i * 3).ColumnWidth = 3
Columns(i * 3 + 1).ColumnWidth = 3
Columns(i * 3 + 2).ColumnWidth = 3
Set riqirange1 = Range(Cells(3, i * 3), Cells(33, i * 3))
Set riqirange2 = Range(Cells(3, i * 3 + 1), Cells(33, i * 3 + 1))
Set riqirange3 = Range(Cells(3, i * 3 + 2), Cells(33, i * 3 + 2))
riqirange0.Copy riqirange1
riqirange2.Select
With Selection
    .Interior.ThemeColor = xlThemeColorAccent3
    .Interior.TintAndShade = 0.8
End With
riqirange3.Select
With Selection
    .Interior.Color = 13434879
    .Interior.TintAndShade = 0.8
End With
b = b + 3
Next i
Range(Cells(1, 1), Cells(33, i * 3 - 1)).Select
With Selection
    .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).ColorIndex = 0
        .Borders(xlEdgeTop).TintAndShade = 0
        .Borders(xlEdgeTop).Weight = xlThin
    .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).ColorIndex = 0
        .Borders(xlEdgeBottom).TintAndShade = 0
        .Borders(xlEdgeBottom).Weight = xlThin
    .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).ColorIndex = 0
        .Borders(xlEdgeRight).TintAndShade = 0
        .Borders(xlEdgeRight).Weight = xlThin
    .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).ColorIndex = 0
        .Borders(xlEdgeLeft).TintAndShade = 0
        .Borders(xlEdgeLeft).Weight = xlThin
    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).ColorIndex = 0
        .Borders(xlInsideHorizontal).TintAndShade = 0
        .Borders(xlInsideHorizontal).Weight = xlThin
    .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideVertical).ColorIndex = 0
        .Borders(xlInsideVertical).TintAndShade = 0
        .Borders(xlInsideVertical).Weight = xlThin
End With
Range(Cells(1, 1), Cells(1, i * 3 - 1)).Select
With Selection
    .Merge
    .Font.Size = 24
    .Font.Name = "微软雅黑"
    .Value = Year(Now) & "年" & Month(Now) & "月台班记录统计"
End With
Cells(34, 3).Select
ActiveCell.FormulaR1C1 = "=COUNT(R3C[1]:R33C[1])"
Range("C34:E34").Select
Selection.AutoFill Destination:=Range("C34:DF34"), Type:=xlFillDefault
Range(Cells(37, 1), Cells(37, 5)).Select
With Selection
    .Merge
    .Font.Size = 12
    .Font.Name = "微软雅黑"
    .FormulaR1C1 = Year(Now) & "年" & Month(Now) & "月油耗统计"
End With
Cells(38, 4).Select
ActiveCell.FormulaR1C1 = "=IFERROR(INT(RC[-2]/RC[-1]),0)"
Range(Cells(38, 4), Cells(i + 37, 4)).FillDown
End Sub
Sub 日期列(qishi As Integer)
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
