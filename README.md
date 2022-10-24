# excel-filter
python script used to filter .xls or .xlsx file
python小工具，wps和excel筛选功能限制较多，工具可用来搜索文件比较大的.xls或者.xlsx，速度比较快。
使用步骤：
1. 需要安装两个packages
`pip install xlwt`
`pip install xlrd`

2. delByName: 根据关键词和列去筛选对应的行，只有这一行指定的列有该关键词，该列就会被删除
col_num--要匹配关键词的列
sheet_num--文件中的sheet的index，如果只有一个填0
source--源文件
target--筛选后新生成的文件
d--关键词list

3. selByList: 根据关键词和列去筛选对应的行，只有这一行有指定的列有该关键词，该列就会被保留
col_num--要匹配关键词的列
sheet_num--文件中的sheet的index，如果只有一个填0
source--源文件
target--筛选后新生成的文件
d--关键词list
