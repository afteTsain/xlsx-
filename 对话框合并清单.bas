Option Explicit

Sub 对话框合并()
    
    Dim bt As Range, r As Long, c As Long, lastcell As Range
    Dim wt As Worksheet
    
    r = 4    'r 是表头的行数
    c = 41    'c 是表头的列数
    
    Dim FileName As String, sht As Worksheet, wb As Workbook
    Dim Erow As Long, fn As String, arr As Variant
    
    Dim Wj_Odj, Wjlx As String, i As Integer '对话框返回文件列表；可选文件类型；循环计数器
    
    
    On Error Resume Next  '容错语句，防止用户取消选择对话框
    
    
     Rem 可选文件类型
    Wjlx = "Excel 97-03版文件(*.Xls),*.Xls,Excel 07版文件(*.Xlsx),*.Xlsx,Excel文件(*.Xl*),*.Xl*"
    Rem 启用打开对话框
    Wj_Odj = Application.GetOpenFilename(FileFilter:=Wjlx, FilterIndex:=3, Title:="打开", MultiSelect:=True)
    If Err.Number <> 0 Then Exit Sub  '检测到用户取消选择对话框则退出过程
       
       Rem 关闭刷新及提示
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set wt = ThisWorkbook.Worksheets(1)    '将汇总表赋给变量wt
    Set lastcell = wt.UsedRange.End(xlDown) '获取最后非空单元格
    Rem 清除汇总表中原表数据,只保留表头
    wt.Rows(2 & ":" & lastcell.Row).ClearContents
    
    Rem 循环获取选中对象的具体文件信息
    For i = 1 To UBound(Wj_Odj)
         Erow = wt.Range("A1").CurrentRegion.Rows.Count + 1     ' 取得汇总表中第一条空行行号
        Set wb = GetObject(Wj_Odj(i)) '使用GetObject函数打开指定文件并赋值给Wb变量
        Set sht = wb.Worksheets(1) ' 将要汇总的工作表赋给变量sht
        ' 将工作表中要汇总的记录保存在数组arr里
        arr = sht.Range(sht.Cells(r + 1, "A"), sht.Cells(1048576, "B").End(xlUp).Offset(0, c - 2))
        ' 将数组arr 中的数据写入工作表
        wt.Cells(Erow, "A").Resize(UBound(arr, 1), UBound(arr, 2)) = arr
        wb.Close False
   Next i
    

    wt.ListObjects("表1").Resize Range("$A$1:$AO$" & wt.UsedRange.End(xlDown).Row)
    ThisWorkbook.Worksheets(2).PivotTables("数据透视表1").PivotCache.Refresh
    
    Sheets("Sheet4").Select
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub

