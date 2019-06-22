'获取单元格中的字符串（正则匹配）:四位数字+一位字母，这是站点名称的命名规范。
Function GetStr(rng As Range)
    With CreateObject("VBscript.regexp")
        .Global = True
        .Pattern = "[0-9]{4}[A-Z]"    '表达式
        If .Execute(rng).Count = 0 Then
            GetStr = ""
        Else
            GetStr = .Execute(rng)(0)
        End If
    End With
End Function
