# Excel-VBA-Functions.xlam
包含若干个函数或模块。具体如下：

Function 1_Combine cells using comma.ba </br>
需求：查询多个电信基站的信息，输入是 数字ID以英文逗号‘,’间隔。
实现：Excel自定义函数，将选中的单元格内容连接，以英文逗号‘,’间隔。

Function 2_Get str using regular.bas
需求：获取字符串中的电信基站名称，该名称是四位数字+1位大写字母。
实现：Excel自定义函数，将选中的单元格内容，正则表达式匹配。

Function 3.1_Setting perspective.bas
需求：KPI透视图，因为字段单位、取值范围不同，会引起透视图大小改变。
实现：对透视图大小、轴位置等进行强制设定。

Function 3.2_Copy perspective to Word.bas
需求：KPI透视图，通常需要对几十个指标绘图，并复制粘贴到邮件。
实现：手动绘制出第一个字段（指标）的透视图，后续依次自动切换到其它指标，并将所有图片复制后粘贴到新建Word文档。最后只需一次将Word中的图片复制到邮件中。

Function 4_Reverse string.bas
实现：自定义函数，将单元格中的字符串反向。

Function 5_Create a SQL module.bas
需求：使用SQL语句对Excel表格UPDATE，往往需要新建VBA模块，指定工作表等准备工作，这些代码不太好记忆。需要将这类工作自动化。
实现：在激活的工作簿中插入一个模块，并写入代码。用户只需要在生成的代码块中书写相应字段的SQL进行UPDATE即可。
