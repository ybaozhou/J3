Private Sub CommandButton1_Click() '导入外部数据
    Application.ScreenUpdating = False    '关闭屏幕刷新
    On Error Resume Next    '忽略错误继续执行VBA代码,避免出现错误消息
    Dim Cnn As Object, Sql As String '参数声明,定义数据库连接和SQL语句
    Set Cnn = CreateObject("Adodb.Connection") '创建数据库连接
    Cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties='Excel 8.0;imex=1';data Source=" & ThisWorkbook.Path & "\名单.xls" '将EXCEL文件作为数据库连接，实际并不打开EXCEL
    Sql = "select 班组,姓名,车号,考核 from [Sheet1$] "
    [A2:D1000].ClearContents '清空数据
    [A2].CopyFromRecordset Cnn.Execute(Sql) '存放数据库数据
    Cnn.Close '关闭数据库连接
    Set Cnn = Nothing '将CNN从内存中删除
    Application.ScreenUpdating = True
End Sub
Private Sub CommandButton2_Click() '导入本簿数据
    Application.ScreenUpdating = False    '关闭屏幕刷新
    On Error Resume Next    '忽略错误继续执行VBA代码,避免出现错误消息
    Dim Cnn As Object, Sql As String '参数声明,定义数据库连接和SQL语句
    Set Cnn = CreateObject("Adodb.Connection") '创建数据库连接
    Cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties='Excel 8.0;imex=1';data Source=" & ThisWorkbook.FullName
    Sql = "select 班组,姓名,车号,考核 from [名单$] "
    [A2:D1000].ClearContents '清空数据
    [A2].CopyFromRecordset Cnn.Execute(Sql) '存放数据库数据
    Cnn.Close '关闭数据库连接
    Set Cnn = Nothing '将CNN从内存中删除
    Application.ScreenUpdating = True
End Sub
