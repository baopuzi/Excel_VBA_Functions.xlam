'在激活的工作簿中插入一个模块，并写入代码，代码示例见底部注释块
Sub Create_SQL_Sub()
    Dim module_name
    module_time = "SQL" + Format(Date, "yyyymmdd") & Format(Time, "hhmmss")
    Application.ScreenUpdating = False
    wb_name = ActiveWorkbook.FullName
    sheet_name = ActiveSheet.name
    ActiveWorkbook.VBProject.VBComponents.Add(1).name = module_time
    ActiveWorkbook.VBProject.VBComponents(module_time).CodeModule.InsertLines 1, "Sub " & module_time & "()"
    ActiveWorkbook.VBProject.VBComponents(module_time).CodeModule.InsertLines 2, "    Set Conn = CreateObject(""ADODB.Connection"")"
    ActiveWorkbook.VBProject.VBComponents(module_time).CodeModule.InsertLines 3, "    Conn.Open ""dsn=excel files;dbq=""&" & """" & wb_name & """"
    ActiveWorkbook.VBProject.VBComponents(module_time).CodeModule.InsertLines 4, ""
    ActiveWorkbook.VBProject.VBComponents(module_time).CodeModule.InsertLines 5, "    'SQL = ""select * from [" & sheet_name & "$]" & """"
    ActiveWorkbook.VBProject.VBComponents(module_time).CodeModule.InsertLines 6, "    'Sheets(""" & sheet_name & """).[M2].CopyFromRecordset Conn.Execute(SQL)"
    ActiveWorkbook.VBProject.VBComponents(module_time).CodeModule.InsertLines 7, ""
    ActiveWorkbook.VBProject.VBComponents(module_time).CodeModule.InsertLines 8, ""
    ActiveWorkbook.VBProject.VBComponents(module_time).CodeModule.InsertLines 9, "    Sql1 = ""Update [LoadMNGCell$] set prbLBExeThrdZUl = '65',prbLBExeThrdZDl='65',intraNeighborLoadRelaThrdUl='15',intraNeighborLoadRelaThrdDl='15'  where description like 'cellLocalId=10%'"""
    ActiveWorkbook.VBProject.VBComponents(module_time).CodeModule.InsertLines 10, "    Sql2 = ""Update [LoadMNGCell$] set prbLBExeThrdZUl = '30',prbLBExeThrdZDl='30',intraNeighborLoadRelaThrdUl='0',intraNeighborLoadRelaThrdDl='0'  where description like 'cellLocalId=20%'"""
    ActiveWorkbook.VBProject.VBComponents(module_time).CodeModule.InsertLines 11, "    Conn.Execute (Sql1)"
    ActiveWorkbook.VBProject.VBComponents(module_time).CodeModule.InsertLines 12, "    Conn.Execute (Sql2)"
    ActiveWorkbook.VBProject.VBComponents(module_time).CodeModule.InsertLines 13, ""
    ActiveWorkbook.VBProject.VBComponents(module_time).CodeModule.InsertLines 14, "    Conn.Close: Set Conn = Nothing"
    ActiveWorkbook.VBProject.VBComponents(module_time).CodeModule.InsertLines 15, "end sub"
End Sub

'以下是点击菜单按钮后生成的示例（Alt+F11）
'Sub SQL20190701131820()
'    Set Conn = CreateObject("ADODB.Connection")
'    Conn.Open "dsn=excel files;dbq=" & "C:\Users\MQ\Desktop\下载实例\欺诈检测_data.csv"
'
'    'SQL = "select * from [欺诈检测_data$]"
'    'Sheets("欺诈检测_data").[M2].CopyFromRecordset Conn.Execute(SQL)
'
'
'    Sql1 = "Update [LoadMNGCell$] set prbLBExeThrdZUl = '65',prbLBExeThrdZDl='65',intraNeighborLoadRelaThrdUl='15',intraNeighborLoadRelaThrdDl='15'  where description like 'cellLocalId=10%'"
'    Sql2 = "Update [LoadMNGCell$] set prbLBExeThrdZUl = '30',prbLBExeThrdZDl='30',intraNeighborLoadRelaThrdUl='0',intraNeighborLoadRelaThrdDl='0'  where description like 'cellLocalId=20%'"
'    Conn.Execute (Sql1)
'    Conn.Execute (Sql2)
'
'    Conn.Close: Set Conn = Nothing
'End Sub


