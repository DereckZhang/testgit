Sub Test4()
    Dim Conn As Object, Rst As Object
    Dim strConn As String, strSQL As String
    Dim i As Integer, PathStr As String
    Set Conn = CreateObject("ADODB.Connection")
    Set Rst = CreateObject("ADODB.Recordset")
    PathStr = ThisWorkbook.Path & "\成交1.xls"   '设置工作簿的完整路径和名称
    Select Case Application.Version * 1    '设置连接字符串,根据版本创建连接
    Case Is <= 11
        strConn = "Provider=Microsoft.Jet.Oledb.4.0;Extended Properties=excel 8.0;Data source=" & PathStr
    Case Is >= 12
        strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & PathStr & ";Extended Properties=""Excel 12.0;HDR=YES; IMEX=1"";"""
    End Select
    '设置SQL查询语句
    strSQL = "select * from [成交1$]"
    Conn.Open strConn    '打开数据库链接
    Set Rst = Conn.Execute(strSQL)    '执行查询，并将结果输出到记录集对象
    With Sheet1
        .Cells.Clear
        For i = 0 To Rst.Fields.Count - 1    '填写标题
            .Cells(1, i + 1) = Rst.Fields(i).Name
        Next i
        .Range("A2").CopyFromRecordset Rst
        .Cells.EntireColumn.AutoFit  '自动调整列宽
        .Cells.EntireColumn.AutoFit  '自动调整列宽
    End With
    Rst.Close    '关闭数据库连接
    Conn.Close
    Set Conn = Nothing
    Set Rst = Nothing
End Sub
add a new line
try for dev