'**********************************************************************************************
'       总表模块
'**********************************************************************************************
'////////////记录子模块\\\\\\\\\\\\\\
'总表按钮
Sub 总表生成()
Dim i As Integer
Dim chepai As String
Dim riqi As Date
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
Sub 每日油耗()
Dim i As Integer
For i = 36 To 72
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
'每月台班数
chepai = Right(Cells(i, 1), 5)
If Not Sheet99.Rows(2).Find(chepai) Is Nothing Then
    chewei = Sheet99.Rows(2).Find(chepai).Column
    Cells(i, 4) = Cells(34, chewei)
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

'**********************************************************************************************
'       单元格格式模块
'**********************************************************************************************
'按钮
Sub 表格格式初始化()
Dim i As Integer
For i = 1 To Sheets.Count - 1
With Sheets(i).Activate
    Cells(1, 1).Select
    ActiveWindow.View = xlPageBreakPreview
    'ActiveWindow.View = xlPageLayoutView
    ActiveWindow.Zoom = 100
    Columns("B:S").ColumnWidth = 3.35
    Columns("A").ColumnWidth = 8.38
    Rows("1:1").RowHeight = 25
    Rows("2:51").RowHeight = 15
    Range("A1:S51").Font.Name = "宋体"
    Range("A1:S51").NumberFormatLocal = "G/通用格式"
    Range("A1:S51").Cells.Replace What:=" ", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Rows("48:51").RowHeight = 15
    Cells(48, "A").Orientation = xlVertical
    Cells(5, "A").Orientation = xlVertical
    '页面修改
        LeftHeader = ""
        CenterHeader = ""
        RightHeader = ""
        LeftFooter = ""
        CenterFooter = ""
        RightFooter = ""
        LeftMargin = Application.InchesToPoints(0.708661417322835)
        RightMargin = Application.InchesToPoints(0.708661417322835)
        TopMargin = Application.InchesToPoints(0.354330708661417)
        BottomMargin = Application.InchesToPoints(0.275590551181102)
        HeaderMargin = Application.InchesToPoints(0.31496062992126)
        FooterMargin = Application.InchesToPoints(0.31496062992126)
        PrintHeadings = False
        PrintGridlines = False
        PrintComments = xlPrintNoComments
        PrintQuality = 600
        CenterHorizontally = False
        CenterVertically = False
        Orientation = xlPortrait
        Draft = False
        PaperSize = xlPaperA4
        FirstPageNumber = xlAutomatic
        Order = xlDownThenOver
        BlackAndWhite = False
        Zoom = 100
        PrintErrors = xlPrintErrorsDisplayed
        OddAndEvenPagesHeaderFooter = False
        DifferentFirstPageHeaderFooter = False
        ScaleWithDocHeaderFooter = True
        AlignMarginsHeaderFooter = True
'        EvenPage.LeftHeader.Text = ""
'        EvenPage.CenterHeader.Text = ""
'        EvenPage.RightHeader.Text = ""
'        EvenPage.LeftFooter.Text = ""
'        EvenPage.CenterFooter.Text = ""
'        EvenPage.RightFooter.Text = ""
'        FirstPage.LeftHeader.Text = ""
'        FirstPage.CenterHeader.Text = ""
'        FirstPage.RightHeader.Text = ""
'        FirstPage.LeftFooter.Text = ""
'        FirstPage.CenterFooter.Text = ""
'        FirstPage.RightFooter.Text = ""
End With
MsgBox 1
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
