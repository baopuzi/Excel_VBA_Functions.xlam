'自定义函数，将多个单元格内容合并为字符串，并以逗号隔开。
Function Com(WorkRng As Range, Optional Sign As String = ",") As String
Dim rng As Range
Dim OutStr As String
For Each rng In WorkRng
'判断单元格内容，内容为英文逗号','或者'#N/A'则忽略。
If rng.Text <> "," And rng.Text <> "#N/A" And Not rng.Text = "" Then
OutStr = OutStr & rng.Text & Sign
End If
Next
Com = Left(OutStr, Len(OutStr) - 1)
End Function
