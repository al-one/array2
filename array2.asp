<%
'ASP键值对数组
'依赖:EasyIDE ASP Framework - http://www.2n.hk/view/i-EasyIDE-ASP-Framework.html
'作者:Alone
'邮箱:Alone@an56.net
'主页:http://www.2n.hk/
'时间:2014-03-15
'说明:您可以免费使用此代码，但请在使用过程中保留上述信息。


function array2(arr,byval k,byval v)
  dim n
  if not array2_is(arr) then new_array2 arr
  if isobject(k) then
    if lcase(typename(k)) = "field" then k = k.value
  end if
  if isobject(v) then
    if lcase(typename(v)) = "field" then v = v.value
  end if
  if not arr.Exists(k) then
    if inull(k) then
      n = array2_max_key(arr)
      n = iif(n < 0,0,n + 1)
      k = n
    end if
    arr.Add k,v
  else
    if isobject(v) then
      set arr.Item(k) = v
    else
      arr.Item(k) = v
    end if
  end if
  set array2 = arr
end function

function new_array2(arr)
  set arr = nothing
  set arr = CreateObject("Scripting.Dictionary")
  set new_array2 = arr
end function

function array2_read(arr,k)
  if not array2_exists(arr,k) then exit function
  if isobject(arr(k)) then
    set array2_read = arr(k)
  else
    array2_read = arr(k)
  end if
end function

function array2_key(arr)
  if not array2_is(arr) then exit function
  array2_key = arr.Keys
end function

function array2_val(arr)
  if not array2_is(arr) then exit function
  array2_val = arr.Items
end function

function array2_ubound(arr)
  dim n : n = -1
  if array2_is(arr) then n = ubound(arr.Keys)
  array2_ubound = n
end function

function array2_max(byval arr)
  dim i,n,m : n = -1
  for i = 0 to ubound(arr)
    if isnumeric(arr(i)) then
      m = clng(arr(i))
      if m > n then n = m
    end if
  next
  array2_max = n
end function

function array2_max_key(byval arr)
  array2_max_key = array2_max(arr.Keys)
end function

function array2_max_val(byval arr)
  array2_max_val = array2_max(arr.Items)
end function

function array2_exists(byval arr,k)
  dim tmp : tmp = false
  if array2_is(arr) then
    tmp = arr.Exists(k)
  end if
  array2_exists = tmp
end function

function array2_is(byval arr)
  dim tmp : tmp = false
  if isobject(arr) then
    tmp = true
  end if
  array2_is = tmp
end function

function array2_is2(byval arr)
  On Error Resume Next
  Err.Clear
  dim tmp : tmp = false
  if isobject(arr) then
    tmp = arr.Keys
    tmp = arr.Items
    if Err.number = 0 then
      tmp = true
    else
      tmp = false
    end if
  end if
  array2_is = tmp
end function

function array2_clone(byval arr,byref arr2)
  set arr2 = arr
  set array2_clone = arr2
end function

function array2_rs(arr,rs)
  dim i,j,arr2
  if not array2_is(arr) then new_array2 arr
  if not isobject(rs) then exit function
  i = 0
  do while not rs.eof
    new_array2 arr2
    for j = 0 to rs.fields.count - 1
      array2 arr2,rs(j).name,rs(j).value
    next
    array2 arr,i,arr2
    rs.movenext
    i = i + 1
  loop
  if not rs.bof then rs.movefirst
  set array2_rs = arr
end function

function array2_match(byval str,pat,arr)
  if not array2_is(arr) then new_array2 arr
  if inull(str) or inull(pat) then exit function
  dim reg,mat,m,i,j,arr2
  set reg = new RegExp
  with reg
    .IgnoreCase = true
    .Global = true
    .Pattern = pat
    set mat = .Execute(str)
  end with
  set reg = nothing
  i = 0
  for each m in mat
    new_array2 arr2
    array2 arr2,0,m.value
    for j = 0 to m.SubMatches.count - 1
      array2 arr2,j + 1,m.SubMatches(j)
    next
    array2 arr,i,arr2
    i = i + 1
  next
  set mat = nothing
  set array2_match = arr
end function

function array2_json_encode(byval arr)
  array2_json_encode = array2_json_encode_obj(arr).jsString
end function

function array2_json_encode_obj(byval arr)
  dim json,k
  if not array2_is(arr) then exit function
  set json = jsObject()
  for each k in arr
    if array2_is(arr(k)) then
      set json(k) = array2_json_encode_obj(arr(k))
    elseif isarray(arr(k)) then
      json(k) = json.toJSON(arr(k))
    else
      json(k) = arr(k)
    end if
  next
  set array2_json_encode_obj = json
  set json = nothing
end function

function array2_dump(byval arr)
  array2_dump = array2_idump(arr,0)
end function

function array2_idump(byval arr,byval s)
  dim k,x,str,pre
  str = vbnewline
  pre = ""
  s = iif(isnumeric(s),clng(s),0)
  for x = 1 to s
    pre = pre & "  "
  next
  if array2_is(arr) then
    for each k in arr
      str = str & pre & k & " => "
      if array2_is(arr(k)) then
        s = s + 1
        str = str & "(" & array2_idump(arr(k),x) & pre & ")"
      elseif isarray(arr(k)) then
        str = str & "array(" & ubound(arr(k)) & "),"
      else
        str = str & arr(k) & ","
      end if
      str = str & vbnewline
    next
  end if
  array2_idump = str
end function
%>
