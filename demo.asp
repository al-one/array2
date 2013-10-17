<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="easyide.asp"-->
<!--#include file="array2.asp"-->
<%
'arr = array()

arr = array2(arr,1,"a")
arr = array2(arr,2,"b")

foreach_arr arr
echo vbnewline & "------------" & ubound(arr) & vbnewline
foreach_arr arr(0)
echo vbnewline & "------------" & ubound(arr(0)) & vbnewline
foreach_arr arr(1)
echo vbnewline & "------------" & ubound(arr(1)) & vbnewline

array2 arr,3,"c"
array2 arr,4,"d"
array2 arr,2,"bb"

foreach_arr arr
echo vbnewline & "------------" & ubound(arr) & vbnewline
foreach_arr arr(0)
echo vbnewline & "------------" & ubound(arr(0)) & vbnewline
foreach_arr arr(1)
echo vbnewline & "------------" & ubound(arr(1)) & vbnewline

echo vbnewline & "++++++++++++" & vbnewline & vbnewline

array2 arr3,111,"aaa"
array2 arr3,222,"bbb"
array2 arr3,333,arr

array2 arr2,null,arr
array2 arr2,"2",arr3
array2 arr2,null,arr
array2 arr2,null,3
array2 arr2,null,4
array2 arr2,null,array(1,2,3,4,5)
array2 arr2,8,array(6,7,8,9,0)

foreach_arr arr2
echo vbnewline & "------------" & ubound(arr2(0)) & vbnewline & array2_read(arr2,5)
echo vbnewline & "" & ubound(array()) & vbnewline

echo array2_dump(arr2,0)


sub foreach_arr(arr)
  if not isarray(arr) then echo arr
  for i = 0 to ubound(arr)
    if isarray(arr(i)) then
      foreach_arr(arr(i))
      echo vbnewline
    else
      echo i & ":" & arr(i) & vbnewline
    end if
  next
end sub

%>