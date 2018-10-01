Sub xlsx批量清洗()

    Dim FileName As String, wb As Workbook
    Dim Erow As Long, fn As String, FilePath As String
    
    Application.ScreenUpdating = False
    
    FilePath = Dir(ThisWorkbook.Path & "\xlsx", vbDirectory)
    If FilePath = "" Then'判断输出目录是否存在
        MkDir (ThisWorkbook.Path & "\xlsx")
    End If
    Kill ThisWorkbook.Path & "\xlsx\*.xlsx"'删除输出目录下所有的工作簿
    FileName = Dir(ThisWorkbook.Path & "\*.xlsx")
    
    Do While FileName <> ""
        If FileName <> ThisWorkbook.Name Then        ' 判断文件是否是汇总数据的工作簿
                     fn = ThisWorkbook.Path & "\" & FileName     '将第1个要汇总的工作簿名称赋给变量fn
            Set wb = GetObject(fn)        ' 将变量fn 代表的工作簿对象赋给变量wb
            FilePath = ThisWorkbook.Path & "\xlsx\" & FileName
            wb.SaveCopyAs FilePath     
             wb.Close False
            Erow = Erow + 1'工作簿计数器
        End If
        FileName = Dir    ' 用Dir 函数取得其他文件名，并赋给变量
    Loop
    
    Application.ScreenUpdating = True
    MsgBox Erow
End Sub
