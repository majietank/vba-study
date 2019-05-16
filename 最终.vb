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
Dim i As Integer, j As Integer, chewei As Integer
Dim chepai As String
Sheet99.Select
j = Range(Cells(36, 1), Cells(65536, 1)).End(xlDown).Row
For i = 36 To j
chepai = Right(Cells(i, 1), 5)
If Not Sheet99.Rows(2).Find(chepai) Is Nothing Then
    chewei = Sheet99.Rows(2).Find(chepai).Column
    Cells(i, 3) = Cells(34, chewei)
End If
If Cells(i, 3) <> 0 Then
    Cells(i, 4) = CInt(Cells(i, 2) / Cells(i, 3))
End If
Next i
End Sub

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
Sub 复制油耗()
Dim i As Integer, j As Integer, chewei As Integer, mryh As Integer, ii As Integer, by As Integer
Dim chepai As String
Sheet99.Select
j = Range(Cells(36, 1), Cells(65536, 1)).End(xlDown).Row
For i = 36 To j
chepai = Right(Cells(i, 1), 5)
mryh = Cells(i, 4)
If Cells(i, 3) > 0 And Cells(i, 4) > 0 Then
Sheets(chepai).Select
For ii = 16 To 46
     If Cells(ii, 16) <> 0 Then
        Cells(ii, 17) = mryh
    End If
Next ii
If Not Range("S16:S46").Find("*保") Is Nothing Then
    by = Range("S16:S46").Find("*保").Row
    If Cells(by, 15) = "√" Then
        Cells(by, 16) = 8 / 2
        Cells(by, 17) = CInt(mryh / 2)
    End If
End If
Sheet99.Select
End If
Next i
End Sub
Sub 设置保月格式()
Dim i As Integer
Sheet99.Select
For i = 1 To Sheets.Count - 1
    Sheets(i).Select
    If Not Range("S16:S46").Find("*保") Is Nothing Then
        Range("S16:S46").Find("*保").Font.Size = 10
    End If
    Sheet99.Select
Next i
End Sub

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