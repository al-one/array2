<%
'ASP键值对数组
'依赖:EasyIDE ASP Framework - http://www.2n.hk/view/i-EasyIDE-ASP-Framework.html
'作者:Alone
'邮箱:Alone@an56.net
'主页:http://www.al-one.cn/
'时间:2013-12-25
'说明:您可以免费使用此代码，但请在使用过程中保留上述信息。


function array2(arr,byval k,byval v)
  dim n
  if not array2_is(arr) then arr = array(array(),array())
  if not arr_in(arr(0),k) then
    if inull(k) then
      n = array2_max(arr(0))
      n = iif(n < 0,0,n + 1)
      k = n
    end if
    arr(0) = arr_push(arr(0),k)
    arr(1) = arr_push(arr(1),v)
  else
    arr(1)(arr_getindex(arr(0),k)) = v
  end if
  array2 = arr
end function

function array2_read(arr,k)
  dim i
  if not array2_is(arr) then exit function
  i = arr_getindex(arr(0),k)
  if i < 0 then exit function
  array2_read = arr(1)(i)
end function

function array2_key(arr)
  if not array2_is(arr) then exit function
  array2_key = arr(0)
end function

function array2_val(arr)
  if not array2_is(arr) then exit function
  array2_val = arr(1)
end function

function array2_ubound(arr)
  dim n : n = -1
  if array2_is(arr) then n = ubound(arr(0))
  array2_ubound = n
end function

function array2_max(arr)
  dim i,n,m : n = -1
  for i = 0 to ubound(arr)
    if isnumeric(arr(i)) then
      m = clng(arr(i))
      if m > n then n = m
    end if
  next
  array2_max = n
end function

function array2_is(arr)
  dim tmp : tmp = false
  if isarray(arr) then
    if ubound(arr) = 1 then
      if isarray(arr(0)) and isarray(arr(1)) then
        if ubound(arr(0)) = ubound(arr(1)) then tmp = true
      end if
    end if
  end if
  array2_is = tmp
end function

function array2_rs(arr,rs)
  dim i,j,arr2
  if not isobject(rs) then exit function
  do while not rs.eof
    for j = 0 to rs.fields.count - 1
      array2 arr2,rs.fields(j).name,rs(j)
    next
    array2 arr,null,arr2
    rs.movenext
  loop
  rs.movefirst
  array2_rs = arr
end function

function array2_dump(arr,byval s)
  dim i,x,str,pre
  str = vbnewline
  pre = ""
  s = iif(isnumeric(s),clng(s),0)
  for x = 1 to s
    pre = pre & "  "
  next
  for i = 0 to iif(array2_is(arr),array2_ubound(arr),ubound(arr))
    str = str & pre & arr(0)(i) & " => "
    if array2_is(arr(1)(i)) then
      s = s + 1
      str = str & "(" & array2_dump(arr(1)(i),x) & ")"
    elseif isarray(arr(1)(i)) then
      str = str & "array(" & ubound(arr(1)(i)) & "),"
    else
      str = str & arr(1)(i) & ","
    end if
    str = str & vbnewline
  next
  array2_dump = str
end function
%>
