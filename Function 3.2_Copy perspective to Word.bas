Sub copy_to_word()
'创建word文件，依次复制粘贴图片，保存
'office 2003, VBA工具/引用中要勾选Microsoft Word 11.0 Object Library
'office 2007, VBA工具/引用中要勾选Microsoft Word 12.0 Object Library
'...

'如果不存在透视图则直接退出执行程序
If ActiveSheet.PivotTables.Count = 0 Then
MsgBox "PivotTable does not exist"
Exit Sub
End If
        
'获取第一个数据透视图的名称，一般叫做“数据透视图1”
Dim pivotTable_name 
pivotTable_name = ActiveSheet.PivotTables(1).name

'检查透视图的第一张图片是否已经设置（轴、图例、值）
'检查轴设置
If ActiveSheet.PivotTables(pivotTable_name).RowFields.Count = 0 Then
MsgBox "Axis is null！"
Exit Sub
End If

'检查图例设置            
If ActiveSheet.PivotTables(pivotTable_name).ColumnFields.Count = 0 Then
MsgBox "Legend is null！"
Exit Sub
End If
                
'检查值设置               
If ActiveSheet.PivotTables(pivotTable_name).DataFields.Count = 0 Then
MsgBox "Value is null！"
Exit Sub
End If

'存在图片，则获取获取当前这个透视图的字段取值方式（求和、计数还是求平均）                   
value_type = ActiveSheet.PivotTables(pivotTable_name).DataFields(1).Function





Dim name, N, i, defpath, fileName, arr(), sheetname
'获取当前sheet
sheetname = ActiveSheet.name 


'fileName1 = Split(ActiveWorkbook.name, ".")(0) '利用split函数分割文件扩展名，但是文件名中有 . 符号就有问题，要合并处理分割后得到的数组

'获取文件名（不带扩展名），目的是创建的word文件相同命名                        
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
fileName = CStr(fso.getbasename(ActiveWorkbook.name))
Set fso = Nothing

'下列部分是测试代码，目的是识别透视图的字段设置
'Dim pivotitem, pivotitem_col, pivotitem1, pivotitem2, pivotitem3
'pivotitem1 = ActiveWorkbook.Sheets("Sheet2").PivotTables("数据透视表1").PivotFields.value '返回所有字段中的单个值
'pivotitem2 = ActiveWorkbook.Sheets("Sheet2").PivotTables("数据透视表1").VisibleFields(1) '返回可见字段中的单个值
'pivotitem3 = ActiveWorkbook.Sheets("Sheet2").PivotTables("数据透视表1").HiddenFields(3) '返回隐藏字段中的单个值
'pivotitem = ActiveWorkbook.Sheets("Sheet2").PivotTables("数据透视表1").HiddenFields.Count'隐藏字段的个数
'pivotitem1 = ActiveWorkbook.Sheets("Sheet2").PivotTables("数据透视表1").ColumnFields.Count '图例
'pivotitem2 = ActiveWorkbook.Sheets("Sheet2").PivotTables("数据透视表1").DataFields(1) '值
'pivotitem2 = ActiveWorkbook.Sheets("Sheet2").PivotTables("数据透视表1").DataFields(1).name '值字段设置
'pivotitem3 = ActiveWorkbook.Sheets("Sheet2").PivotTables("数据透视表1").RowFields(1)'轴
'pivotitem_col = WorksheetFunction.Match(pivotitem2, ActiveWorkbook.Sheets("Sheet2").PivotTables("数据透视表1").PivotField, 0)


'找出设置的“值”是第n个字段
With ActiveSheet.PivotTables(pivotTable_name)
name = .DataFields(1).name
m = .PivotFields.Count
ReDim arr(1 To m)
    For i = 1 To m
        arr(i) = .PivotFields(i).name
        If Right(name, Len(arr(i))) = arr(i) Then
                N = i
                value_type_name = Left(name, Len(name) - Len(arr(i)))
        End If
    Next
End With


ActiveSheet.ChartObjects(1).Activate '选中第一个图
ActiveChart.ChartArea.Copy '复制透视图图片，运行此代码前要先点击选中数据透视图；目前未解决这个bug

Dim WordApp As Word.Application '定义变量
Set WordApp = CreateObject("Word.Application") '生成WORD对象
WordApp.Documents.Add '新建文件

    
  
    defpath = ActiveWorkbook.Path '获取Excel文件路径
    fn$ = defpath & "\" & fileName & "_" & Format(Date, "yyyymmdd") & Format(Time, "hhmmss") & ".docx" '生成文件名
    WordApp.ActiveDocument.SaveAs fn$ '按照以上生成的文件名保存文件
    
    If Documents.Count > 0 Then
    For Each doc In Documents
    'MsgBox fn$
    'MsgBox doc.Path & "\" & doc.name
    Next doc
    Else
    MsgBox "Please test again！"
    End If
    
    '设置Word页面的宽和高，为了使从Word复制到Zmail的图片尽量清晰，所以要保证图片的宽和高
    With WordApp.Selection.PageSetup
         .PageWidth = CentimetersToPoints(36.4)
         .PageHeight = CentimetersToPoints(25.7)
    End With
      
   'WordApp.Visible = True '设置生成的Word文件在运行过程中是否可见
    WordApp.Selection.PasteSpecial DataType:=wdPasteBitmap  '将Excel中的数据透视图以图片格式粘贴到Word
    
    '依次将后续透视图复制粘贴到Word
    If N < m Then
    For i = N + 1 To m
    ActiveSheet.PivotTables(pivotTable_name).PivotFields( _
        value_type_name & arr(i - 1)).Orientation = xlHidden
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields(arr(i)), _
        value_type_name & arr(i), value_type
    ActiveChart.ChartArea.Copy
'   WordApp.Selection.PasteSpecial DataType:=wdPasteBitmap  '粘贴
    WordApp.Selection.PasteAndFormat (wdChartPicture)
    Next
    End If
   
   
    '将数据透视图恢复成初始设置，显示第一个值
    ActiveSheet.PivotTables(pivotTable_name).PivotFields( _
        value_type_name & arr(m)).Orientation = xlHidden
     
  '  ActiveSheet.PivotTables(pivotTable_name).AddDataField ActiveSheet.PivotTables("数据透视表2" _
  '  ).PivotFields(arr(n)), _
  '      "求和项:" & arr(n), xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields(arr(N)), _
        value_type_name & arr(N), value_type
        
     


WordApp.ActiveDocument.Save '保存Word文件
WordApp.Quit '退出
Set WordApp = Nothing '取消变量
 
'Exit Sub '可能遇到未知错误，程序要结束执行
'ErrHandler:
    'MsgBox "遇到未知错误，请保存并关闭Excel后重试！"
    'Set WordApp = Nothing '取消变量
End Sub
