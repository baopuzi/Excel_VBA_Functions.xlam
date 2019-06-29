'获取单元格中的字符串（正则匹配）:四位数字+一位字母，这是站点名称的命名规范。
Function GetStr(rng As Range)
    With CreateObject("VBscript.regexp")
        .Global = True
        .Pattern = "[0-9]{4}[A-Z]"    '正则表达式，四位数字和一个大写字母
        If .Execute(rng).Count = 0 Then
            GetStr = ""               '如果没有匹配到，以空字符串代替，需检查原字符串是否错误
        Else
            GetStr = .Execute(rng)(0) '如果匹配到，则取第一次匹配到的结果
        End If
    End With
End Function
