<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="easyide.asp"-->
<!--#include file="array2.asp"-->
<%
Server.ScriptTimeout = 999999
Response.Codepage = 65001
Response.Charset  = "UTF-8"


'性能测试
'------------------------------------
debug_stime = timer
echo "<textarea>"
for i = 0 to 9999
  echo i & vbnewline
next
echo "</textarea><br>"
echo "循环10000次：" & debug_time(debug_stime)


echo vbnewline & "<br><br><br>" & vbnewline


debug_stime = timer
n = 0
for i = 1 to 10000
  n = n + i
next
echo "1加到10000 = " & n & "：" & debug_time(debug_stime)


echo vbnewline & "<br><br><br>" & vbnewline


debug_stime = timer
for i = 1 to 1000
  array2 arr,i,"v" & i
next
echo "array2新增1000个元素：" & debug_time(debug_stime)


echo vbnewline & "<br><br><br>" & vbnewline


debug_stime = timer
for i = 1 to 100
  array2 arr,i,"vv" & i
next
echo "array2修改100个元素：" & debug_time(debug_stime)


echo vbnewline & "<br><br><br>" & vbnewline


debug_stime = timer
echo "<textarea>" & array2_dump(arr) & "</textarea><br>"
echo "array2遍历1000个元素：" & debug_time(debug_stime)


echo vbnewline & "<br><br><br>" & vbnewline


debug_stime = timer
echo "<textarea>"
for i = 0 to 999
  echo array2_read(arr,i) & vbnewline
next
echo "</textarea><br>"
echo "array2读取1000次：" & debug_time(debug_stime)


echo vbnewline & "<br><br><br>" & vbnewline




'用法示例
'初始化
'------------------------------------
arr = null
'或
new_array2 arr

'新增/修改
'------------------------------------
array2 arr3,null,"x"
array2 arr3,null,"y"
array2 arr3,2,"z"
array2 arr2,null,"AA"
array2 arr2,1,"BB"
array2 arr2,null,arr3
array2 arr,0,"v0"
array2 arr,null,"v1"
array2 arr,null,arr2

'遍历
'------------------------------------
echo "<h3>新增：</h3><pre>" & array2_dump(arr) & "</pre><br>"




function debug_time(debug_stime)
  debug_time = FormatNumber((timer - debug_stime) * 1000,3)
end function
%>
<hr>
<pre>
新增1000/读取1000(毫秒)
Log v2.1:
    7.813/7.813

Log v1.1:
  IIS5.1@WinXP
    1,890.625/265.625
    2,062.500/250.000
    1,796.875/218.750
    2,125.000/312.500
    1,984.375/250.000
    1,843.750/281.250
    1,875.000/218.750
    1,890.625/203.125
    2,000.000/234.375
    1,875.000/218.750

  IIS6.1@Win7:
    703.125/93.750
    667.969/93.750
    746.094/93.750
    671.875/93.750
    703.125/93.750
    734.375/93.750
    750.000/93.750
    734.375/93.750
    687.500/93.750
    796.875/93.750

Log v1.0:
  IIS5.1@WinXP:
    2,843.750/218.750
    2,781.250/218.750
    2,984.375/234.375
    2,921.875/203.125
    2,828.125/203.125
    2,875.000/203.125
    2,875.000/218.750
    2,812.500/250.000
    3,000.000/218.750
    3,156.250/234.375

  IIS6.1@Win7:
    1,187.500/109.375
    1,136.719/93.750
    1,167.969/109.375
    1,125.000/93.750
    1,187.500/156.250
    1,136.719/93.750
    1,140.625/93.750
    1,156.250/113.281
    1,156.250/93.750
    1,218.750/93.750
</pre>