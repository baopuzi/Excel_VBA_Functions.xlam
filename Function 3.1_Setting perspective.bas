Sub set_size()
'
' 设置透视图大小 宏
'
    If ActiveSheet.Shapes.Count <> 0 Then
    With ActiveSheet.Shapes("Chart 1")
    .Height = 320 '设置透视图高
    .Width = 800 '设置透视图宽
    .Top = 0 '设置透视图上边距
    .Left = 0 '设置透视图左边距
    .Placement = xlFreeFloating '设置透视图的大小固定
    End With
    ActiveSheet.ChartObjects("Chart 1").Activate '激活透视图
    With ActiveChart
    .ChartColor = 10
    .Axes(xlCategory).TickLabelPosition = xlLow '设置横轴方位为下方
    .ChartColor = 10 '设置图表的线条颜色，因为多次运行切换字段值后，所有线条会自动变为同一种颜色
    '设置图例位置
    .Legend.Height = 40
    .Legend.Width = 710
    .Legend.Top = 275
    .Legend.Left = 58
    '设置绘图区位置，包含坐标轴
    .PlotArea.Top = 17
    .PlotArea.Left = 10
    .PlotArea.Width = 790
    .PlotArea.Height = 200
    '设置绘图区中的图形区域位置，不包含坐标轴
    .PlotArea.InsideWidth = 740
    .PlotArea.InsideHeight = 170
    .PlotArea.InsideTop = 22
    .PlotArea.InsideLeft = 200 '已经超过PlotArea允许范围
    End With
    Else                       '可能在非透视图情况下点击按键，如果透视图数量为0，表明不存在透视图，退出执行程序。             
    MsgBox "Chart 1 does not exist"  
    Exit Sub
    End If
End Sub
