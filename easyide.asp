<%
' EasyIDE ASP Framework By Alone
' 作者:Alone
' 邮箱:Alone@an56.net
' 时间:2015-05-16
' 说明:此函数库原作者“沉沦”，本人增加和修改了一些函数。
'      您可以免费使用此库，但请在使用过程中保留上述信息。


Const EASYIDE_CHARSET = "utf-8"
Const OBJ_RST  = "ADODB.Recordset"
Const OBJ_CONN = "ADODB.Connection"
Const OBJ_STRM = "ADODB.Stream"
Const OBJ_FSO  = "Scripting.FilesyStemObject"
Const OBJ_XHTP = "MSXML2.XMLHTTP"
Const OBJ_DOM  = "MSXML2.DOMDocument"


'/////////基础操作函数部分

'函数：安全获取参数
function rqs(str)
  rqs = str_safe(Request(str))
end function

'过程：输出字符串[代替Response.Write]
sub echo(str)
  Response.Write(str)
end sub

'过程：结束页面并输出字符串
sub die(str)
  Response.Write(str) : Response.End()
end sub

'过程：返回信息
sub infoback(str)
  if lcase(typename(rs)) = "recordset" then rs.close : set rs = nothing
  Response.ContentType = "text/html"
  die "<script>alert('" & str & "');window.history.back();</script>"
end sub

sub infohref(str,url)
  if lcase(typename(rs)) = "recordset" then rs.close : set rs = nothing
  Response.ContentType = "text/html"
  die "<script>alert('" & str & "');window.location.href='" & url & "';</script>"
end sub

'函数：十进制转二进制
function cbit(byval num)
  dim base64
  set base64 = new base64_class
  num = base64.cbit(num)
  set base64 = nothing
  cbit = num
end function

'函数：二进制转十进制
function cdec(byval num)
  dim base64
  set base64 = new base64_class
  num = base64.cdec(num)
  set base64 = nothing
  cdec = num
end function

'函数：毫秒数转换为时间长度
function ctime(byval num,n)
  dim tmp : tmp = 0
  if not isnumeric(num) then ctime = tmp : exit function
  if inull(n) or not isnumeric(n) then n = 2
  if num >= 1000*60*60*24*30 then
    tmp = round(num/(1000*60*60*24*30),n) & "月"
  elseif num >= 1000*60*60*24 then
    tmp = round(num/(1000*60*60*24),n) & "天"
  elseif num >= 1000*60*60 then
    tmp = round(num/(1000*60*60),n) & "小时"
  elseif num >= 1000*60 then
    tmp = round(num/(1000*60),n) & "分钟"
  elseif num >= 1000 then
    tmp = round(num/1000,n) & "秒"
  else
    tmp = round(num,n) & "毫秒"
  end if
  ctime = tmp
end function

'函数：字节数转换为文件大小
function csize(byval num,n)
  dim tmp : tmp = 0
  if not isnumeric(num) then csize = tmp : exit function
  if inull(n) or not isnumeric(n) then n = 2
  if num >= 1024*1024*1024*1024 then
    tmp = round(num/(1024*1024*1024*1024),n) & "TB"
  elseif num >= 1024*1024*1024 then
    tmp = round(num/(1024*1024*1024),n) & "GB"
  elseif num >= 1024*1024 then
    tmp = round(num/(1024*1024),n) & "MB"
  elseif num >= 1024 then
    tmp = round(num/1024,n) & "KB"
  else
    tmp = round(num,n) & "Byte"
  end if
  csize = tmp
end function

'过程：动态包含文件
sub include(path)
  echo file_eval(path)
end sub

'函数：base64加密
function base64encode(byval str)
  if isnull(str) then exit function
  dim base64
  set base64 = new base64_class
  str = base64.encode(str)
  set base64 = nothing
  base64encode = str
end function

'函数：base64解密
function base64decode(byval str)
  if isnull(str) then exit function
  dim base64
  set base64 = new base64_class
  str = base64.decode(str)
  set base64 = nothing
  base64decode = str
end function

'函数：URL加密
function urlencode(byval str)
  if isnull(str) then exit function
  str = Server.URLEncode(str)
  urlencode = str
end function

'函数：Escape加密
function escape(byval str)
  if isnull(str) then exit function
  dim i,c,a,tmp : tmp = ""
  for i=1 to len(str)
    c = mid(str,i,1)
    a = ascw(c)
    if (a>=48 and a<=57) or (a>=65 and a<=90) or (a>=97 and a<=122) then
      tmp = tmp & c
    elseif instr("@*_+-./",c) > 0 then
      tmp = tmp & c
    elseif a>0 and a<16 then
      tmp = tmp & "%0" & hex(a)
    elseif a>=16 and a<256 then
      tmp = tmp & "%" & hex(a)
    else
      tmp = tmp & "%u" & hex(a)
    end if
  next
  escape = tmp
end function

'函数：Escape解密
function unescape(byval str)
  if isnull(str) then exit function
  dim i,c,tmp : tmp = ""
  for i=1 to len(str)
    c = mid(str,i,1)
    if mid(str,i,2)="%u" and i<=len(str)-5 then
      if isnumeric("&H" & mid(str,i+2,4)) then
        tmp = tmp & chrw(cint("&H" & mid(str,i+2,4)))
        i = i+5
      else
        tmp = tmp & c
      end if
    elseif c="%" and i<=len(str)-2 then
      if isnumeric("&H" & mid(str,i+1,2)) then
        tmp = tmp & chrw(cint("&H" & mid(str,i+1,2)))
        i = i+2
      else
        tmp = tmp & c
      end if
    else
      tmp = tmp & c
    end if
  next
  unescape = tmp
end function

'函数：三元IF
function iif(exp,v1,v2)
  dim tmp : tmp = v2
  if exp then tmp = v1
  iif = tmp
end function

'函数：如果真，则...
function ifi(v1,v2)
  dim tmp
  if ift(v1) then tmp = v1 else tmp = v2
  ifi = tmp
end function

'函数：如果真
function ift(val)
  dim tmp : tmp = true
  if isempty(val) then
    tmp = false
  elseif isnull(val) then
    tmp = false
  elseif isdate(val) then
  elseif isobject(val) then
    if val is nothing then tmp = false
  elseif isarray(val) then
    if ubound(val) < 0 then tmp = false
  else
    if isnumeric(val) then
      if val & "" = "0" then tmp = false
    end if
    select case VarType(val)
      case 8 'string
        if trim(val) = "" then tmp = false else tmp = true
      case 11 'bool
        if not val then tmp = false
    end select
  end if
  ift = tmp
end function

'函数：空值测试
function inull(val)
  dim tmp : tmp = false
  if isnull(val) then
    tmp = true
  elseif isempty(val) then
    tmp = true
  elseif trim(val) = "" then
    tmp = true
  end if
  inull = tmp
end function

'全启变量：客户端IP
dim ip : ip = Request.ServerVariables("REMOTE_ADDR")

'函数：返回客户端真实IP
function realip
  dim tmp : tmp = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
  if inull(tmp) then tmp = Request.ServerVariables("REMOTE_ADDR")
  realip = tmp
end function

'函数：邮件发送[Jamil-Message]
function sendmail(fromname,sendto,subject,body,from,serveraddress,username,password)
  dim jmail, return
  set jmail = Server.CreateObject("JMAIL.Message")
  jmail.silent   = true
  jmail.logging  = true
  jmail.charset  = EASYIDE_CHARSET
  jmail.contenttype = "text/html; charset=" & EASYIDE_CHARSET
  jmail.addrecipient sendto
  jmail.fromname = fromname
  jmail.from     = from
  jmail.mailserverusername = username
  jmail.MailServerPassword = password
  jmail.subject  = subject
  jmail.body     = body
  jmail.priority = 3
  return = jmail.send(serveraddress)
  jmail.close()
  set jmail = nothing
  sendmail = return
end function

'函数：检测组件是否安装
function install(str)
  dim tmp : tmp = false
  dim obj_test
  on error resume next
  err.clear()
  set obj_test = Server.CreateObject(str)
  if err.number = 0 then tmp = true
  set obj_test = nothing
  err.clear()
  install = tmp
end function


'/////////字符串操作函数部分

'函数：转换为整型数值
function str_int(byval str)
  dim max,min
  max = 2147483647
  min = -2147483648
  str = str_dbl(str)
  if str > max then
    str = max
  elseif str < min then
    str = min
  end if
  str_int = int(str)
end function

'函数：转换为双精度数值
function str_dbl(byval str)
  str = trim(str)
  if inull(str) then str = 0
  if not isnumeric(str) then str = 0
  str = cdbl(str)
  str_dbl = str
end function

'函数：正则验证
function str_test(pattern,str)
  dim tmp : tmp = false
  dim reg : set reg = new regexp
  with reg
    .ignorecase = true
    .global = true
    .pattern = pattern
    tmp = .test(str)
  end with
  set reg = nothing
  str_test = tmp
end function

'函数：正则替换[不区分大小写]
function str_replace(pattern,byval str,s)
  if isnull(str) then exit function
  dim tmp : tmp = false
  dim reg : set reg = new regexp
  with reg
    .ignorecase = true
    .global = true
    .pattern = pattern
    tmp = .replace(str,s)
  end with
  set reg = nothing
  str_replace = tmp
end function

'函数：正则替换[区分大小写]
function str_ireplace(pattern,byval str,s)
  if isnull(str) then exit function
  dim tmp : tmp = false
  dim reg : set reg = new regexp
  with reg
    .ignorecase = false
    .global = true
    .pattern = pattern
    tmp = .replace(str,s)
  end with
  set reg = nothing
  str_ireplace = tmp
end function

'函数：执行正则搜索并返回结果集[不区分大小写]
function str_execute(pattern,byval str)
  if isnull(str) then exit function
  dim tmp : tmp = false
  dim reg : set reg = new regexp
  with reg
    .ignorecase = true
    .global = true
    .pattern = pattern
    set tmp = .execute(str)
  end with
  set reg = nothing
  set str_execute = tmp
end function

'函数：执行正则搜索并返回结果集[区分大小写]
function str_iexecute(pattern,byval str)
  if isnull(str) then exit function
  dim tmp : tmp = false
  dim reg : set reg = new regexp
  with reg
    .ignorecase = false
    .global = true
    .pattern = pattern
    set tmp = .execute(str)
  end with
  set reg = nothing
  set str_iexecute = tmp
end function

'函数：精确计算字符串长度
function str_len(byval str)
  str = str_replace("[^\x00-\xff]",str,"@@")
  str_len = len(str)
end function

'函数：截断字串
function str_left(byval str,slen,ext)
  if isnull(str) then exit function
  dim tmp : tmp = "&nbsp;= |&quot;=""|&amp;=&|&lt;=<|&gt;=>|&euro;= |&laquo;=《|&raquo;=》|&hellip;= |&copy;= "
  dim arr, a, v : arr = split(tmp,"|")
  for each v in arr
    a = split(v,"=")
    str = replace(str,a(0),a(1))
  next
  dim i, c, s, n : n = 0 : tmp = ""
  for i = 1 to len(str)
    s = mid(str,i,1)
    c = abs(ascw(s))
    if c > 255 then n = n + 2 else n = n + 1
    tmp = tmp & s
    if n >= slen then exit for
  next
  if tmp = str then ext = ""
  str_left = tmp & ext
end function

'函数：返回可安全地用于SQL操作的字符串
function str_safe(byval str)
  if isnull(str) then exit function
  str = str_isafe(str)
  str = replace(str,"<","&lt;")
  str = replace(str,">","&gt;")
  str = replace(str,"""","&quot;")
  str_safe = str
end function

'函数：SQL关键词过滤 用于获取含HTML标签的内容
function str_isafe(byval str)
  if isnull(str) then exit function
  str = replace(str,"select ","sel&#101;ct ",1,-1,1)
  str = replace(str,"insert ","ins&#101;rt ",1,-1,1)
  str = replace(str,"update ","up&#100;ate ",1,-1,1)
  str = replace(str,"delete ","del&#101;te ",1,-1,1)
  str = replace(str," and"," an&#100; ",1,-1,1)
  str = replace(str,"drop table","dro&#112; table",1,-1,1)
  str = replace(str,"script","&#115;cript")
  str = replace(str,"*","&#42;")
  str = replace(str,"%","&#37;")
  str = replace(str,"'","''")
  str_isafe = str
end function

'函数：替换简单HTML格式字符为控制字符
function str_htmldecode(byval str)
  if isnull(str) then exit function
  str = replace(str,"&nbsp;"," ")
  str = replace(str,"<br />",chr(10))
  str_htmldecode = str
end function

'函数：替换字符串中的控制字符为HTML代码。
function str_htmlencode(byval str)
  if isnull(str) then exit function
  str = replace(str," ","&nbsp;")
  str = replace(str,chr(10),"<br />")
  str_htmlencode = str
end function

'函数：清除HTML标签
function str_htmlclear(byval str)
  if isnull(str) then exit function
  str = replace(str,"&nbsp;"," ")
  str = str_replace("<[^>]+?>",str,"")
  str_htmlclear = str
end function

'函数：清除所有格式及空格 压缩字符串
function str_trim(byval str)
  if isnull(str) then exit function
  str = str_replace("<[^>]+?>|\s+",str,"")
  str = replace(str,"&nbsp;","")
  str_trim = str
end function

'函数：清除左右空白字符
function str_atrim(byval str)
  if isnull(str) then exit function
  str = str_replace("^\s+|\s+$",str,"")
  str_atrim = str
end function

'函数：清除左边空白字符
function str_ltrim(byval str)
  if isnull(str) then exit function
  str = str_replace("^\s+",str,"")
  str_ltrim = str
end function

'函数：清除右边空白字符
function str_rtrim(byval str)
  if isnull(str) then exit function
  str = str_replace("\s+$",str,"")
  str_rtrim = str
end function

'函数：返回一个不重复的随机字串
function str_rnd()
  dim ran_num, tmp
  randomize : ran_num = int((90000 * rnd) + 10000)
  tmp = str_time("ymmddhhiiss",now) & ran_num
  str_rnd = base64encode(tmp)
end function

'函数：返回格式化的时间字串
function str_time(format,byval str)
  if inull(str) or not isdate(str) then exit function
  dim tmp,y,yy,m,mm,d,dd,h,hh,i,ii,s,ss,u,uu
  tmp = format
  y = year(str)   : yy = right("0" & y,2)
  m = month(str)  : mm = right("0" & m,2)
  d = day(str)    : dd = right("0" & d,2)
  h = hour(str)   : hh = right("0" & h,2)
  i = minute(str) : ii = right("0" & i,2)
  s = second(str) : ss = right("0" & s,2)
  u = DateDiff("s","1970-01-01 00:00:00",str) : uu = u * 1000
  tmp = replace(tmp,"yyyy",y,1,-1,1)
  tmp = replace(tmp,"yy",yy,1,-1,1)
  tmp = replace(tmp,"y",y,1,-1,1)
  tmp = replace(tmp,"mm",mm,1,-1,1)
  tmp = replace(tmp,"m",m,1,-1,1)
  tmp = replace(tmp,"dd",dd,1,-1,1)
  tmp = replace(tmp,"d",d,1,-1,1)
  tmp = replace(tmp,"hh",hh,1,-1,1)
  tmp = replace(tmp,"h",h,1,-1,1)
  tmp = replace(tmp,"ii",ii,1,-1,1)
  tmp = replace(tmp,"i",i,1,-1,1)
  tmp = replace(tmp,"ss",ss,1,-1,1)
  tmp = replace(tmp,"s",s,1,-1,1)
  tmp = replace(tmp,"uu",uu,1,-1,1)
  tmp = replace(tmp,"u",u,1,-1,1)
  if instr(lcase(tmp),"w") > 0 then
    dim w,arr
    w = weekday(str)
    arr = array("","日","一","二","三","四","五","六")
    tmp = replace(tmp,"W",arr(w))
    tmp = replace(tmp,"w",w - 1)
  end if
  str_time = tmp
end function

'函数：从字串中分离出远程文件URL
function str_geturl(byval str,ext)
  if isnull(str) then exit function
  dim exts : exts = split(ext,",")
  dim pattern, e, s : pattern = "" : s = ""
  for each e in exts
    pattern = pattern & s & "http://[\S]+?\."&e : s = "|"
  next
  dim matches : set matches = str_execute(pattern,str)
  dim m, urls : urls = "" : s = ""
  for each m in matches
    urls = urls & s & m.value : s = "#"
  next
  str_geturl = split(urls,"#")
end function

'函数：获取URL参数串
function str_query(del)
  dim tmp : tmp = Request.ServerVariables("QUERY_STRING")
  if inull(del) then str_query = tmp : exit function
  tmp = str_param(tmp,del)
  str_query = tmp
end function

'函数：处理URL参数串
function str_param(str,del)
  dim tmp : tmp = str
  if inull(tmp) or inull(del) then str_param = tmp : exit function
  dim arr : arr = split(tmp, "&")
  dim q, a, t : t = "" : tmp = ""
  for each q in arr
    if trim(q) <> "" then
      a = split(q,"=") : if ubound(a) = 0 then arr_push a,""
      if not arr_in(split(del,","),a(0)) then tmp = tmp & t & a(0) & "=" & a(1) : t = "&"
    end if
  next
  str_param = tmp
end function

'函数：预处理URL参数串
function str_iparam(str,key,val)
  dim url,par,tmp
  if str_test("^[^?=&]*\??",str) and not inull(str) then
    tmp = instr(str,"?")
    url = left(str,iif(tmp > 0,tmp - 1,len(str))) & "?"
    par = right(str,len(str) - ifi(tmp,len(str)))
  else
    url = ""
    par = str
  end if
  if inull(key) then
    if not inull(par) and not inull(val) then par = str_param(par,val)
  else
    if not inull(par) then par = str_param(par,key)
    par = par & iif(inull(par),"","&") & key & "=" & val
  end if
  str_iparam = url & par
end function

'函数：字符串加密
function str_encode(byval str)
  if isnull(str) then exit function
  dim base64 : set base64 = new base64_class
  base64.bstr = "ABCDEF1234GHIJKLMnopqrs+tuvwxyz09abcdef!ghijklmNOPQRS5678TUVWXYZ"
  base64.blen = 16
  str = base64.encode(str)
  set base64 = nothing
  str_encode = str
end function

'函数：字符串解密
function str_decode(byval str)
  if isnull(str) then exit function
  dim base64 : set base64 = new base64_class
  base64.bstr = "ABCDEF1234GHIJKLMnopqrs+tuvwxyz09abcdef!ghijklmNOPQRS5678TUVWXYZ"
  base64.blen = 16
  str = base64.decode(str)
  set base64 = nothing
  str_decode = str
end function

'函数：将ASP文件运行结果返回为字串
function str_eval(str)
  dim tmp, a, b, t, matches, m
  tmp = "dim str_eval_htm : str_eval_htm = """"" & vbcrlf
  a = 1
  b = instr(a,str,"<%") + 2
  while b > a + 1
    t = mid(str,a,b - a - 2)
    t = newline_encode(t)
    t = replace(t,"""","""""")
    tmp = tmp & "str_eval_htm = str_eval_htm & """ & t & """" & vbcrlf
    a = instr(b,str,"%\>") + 2
    t = mid(str,b,a - b - 2)
    t = str_replace("^\s*=",t,"str_eval_htm = str_eval_htm & ")'< %=str% >
    t = str_replace("([\s:]*)(?:echo|response.write)(\s+|\()",t,"$1str_eval_htm = str_eval_htm & $2")
    t = str_replace("([\s:]*response.clear(?:\(\s*\))?)",t,"$1 : str_eval_htm = """"")
    t = str_replace("([\s:]*)die(\s+|\()",t,"$1die newline_decode(str_eval_htm) & $2")
    t = str_replace("([\s:]*)(response.end(?:\(\s*\))?)",t,"$1die newline_decode(str_eval_htm) : $2")
    tmp = tmp & t & vbcrlf
    b = instr(a,str,"<%") + 2
  wend
  t = mid(str,a)
  t = newline_encode(t)
  t = replace(t,"""","""""")
  tmp = tmp & "str_eval_htm = str_eval_htm & """ & t & """" & vbcrlf
  'die tmp
  'execute(tmp)
  executeglobal(tmp)
  str_eval_htm = newline_decode(str_eval_htm)
  str_eval = str_eval_htm
end function

function newline_encode(byval str)
  str = replace(str,vbcrlf,"{::vbcrlf}")
  str = replace(str,vbcr,"{::vbcr}")
  newline_encode = str
end function

function newline_decode(byval str)
  str = replace(str,"{::vbcrlf}",vbcrlf)
  str = replace(str,"{::vbcr}",vbcr)
  newline_decode = str
end function


'/////////文件操作函数部分

'函数：获取当前脚本执行文件的文件名
function file_self()
  dim tmp
  tmp = Request.ServerVariables("SCRIPT_NAME")
  tmp = split(tmp,"/")
  file_self = tmp(ubound(tmp))
end function

'函数：获取当前脚本执行文件所在的磁盘目录
function file_dir()
  dim tmp, arr
  tmp = file_self()
  arr = split(Server.MapPath(tmp),"\")
  file_dir = arr(ubound(arr) - 1)
end function

'函数：检测文件/文件夹是否存在
function file_exists(path)
  dim tmp : tmp = false
  dim fso : set fso = Server.CreateObject(OBJ_FSO)
  if fso.fileexists(Server.MapPath(path)) then tmp = true
  if fso.folderexists(Server.MapPath(path)) then tmp = true
  set fso = nothing
  file_exists = tmp
end function

'函数：删除文件/文件夹
function file_delete(path)
  dim tmp : tmp = false
  dim fso : set fso = Server.CreateObject(OBJ_FSO)
  if fso.fileexists(Server.MapPath(path)) then'文件
    fso.deletefile(Server.MapPath(path))
    if not fso.fileexists(Server.MapPath(path)) then tmp = true
  end if
  if fso.folderexists(Server.MapPath(path)) then'文件夹
    fso.deletefolder(Server.MapPath(path))
    if not fso.folderexists(Server.MapPath(path)) then tmp = true
  end if
  set fso = nothing
  file_delete = tmp
end function

'函数：获取文件/文件夹信息
function file_info(path)
  dim tmp(4)
  dim fso : set fso = Server.CreateObject(OBJ_FSO)
  if fso.fileexists(Server.MapPath(path)) then '文件
    dim fl : set fl = fso.getfile(Server.MapPath(path))
    tmp(0) = fl.type'类型
    tmp(1) = fl.attributes'属性
    tmp(2) = csize(fl.size,4)'大小
    tmp(3) = fl.datecreated'创建时间
    tmp(4) = fl.datelastmodified'最后修改时间
  elseif fso.folderexists(Server.MapPath(path)) then '文件夹
    dim fd : set fd = fso.getfolder(Server.MapPath(path))
    tmp(0) = "folder"'类型
    tmp(1) = fd.attributes'属性
    tmp(2) = csize(fd.size,4)'大小
    tmp(3) = fd.datecreated'创建时间
    tmp(4) = fd.datelastmodified'最后修改时间
  end if
  set fso = nothing
  file_info = tmp
end function

'函数：复制文件/文件夹
function file_copy(file_start,file_end,model)
  if model <> 0 and model <> 1 then model = false else model = cbool(model)
  dim tmp : tmp = false
  dim fso : set fso = Server.CreateObject(OBJ_FSO)
  if fso.fileexists(Server.MapPath(file_start)) then '文件
    fso.copyfile Server.MapPath(file_start),Server.MapPath(file_end),model
    if fso.fileexists(Server.MapPath(file_end)) then tmp = true
  end if
  if fso.folderexists(Server.MapPath(file_start)) then '文件夹
    fso.copyfolder Server.MapPath(file_start),Server.MapPath(file_end),model
    if fso.folderexists(Server.MapPath(file_end)) then tmp = true
  end if
  set fso = nothing
  file_copy = tmp
end function

'函数：创建文件夹
function file_create(path,model)
  if model <> 0 and model <> 1 then model = false else model = cbool(model)
  dim tmp : tmp = false
  dim fso : set fso = Server.CreateObject(OBJ_FSO)
  if fso.folderexists(Server.MapPath(path)) then
    if model then fso.deletefolder(Server.MapPath(path)) : fso.createfolder Server.MapPath(path)
  else
    fso.createfolder Server.MapPath(path)
  end if
  if fso.folderexists(Server.MapPath(path)) then tmp = true
  set fso = nothing
  file_create = tmp
end function

'函数：获取指定目录下所有文件及文件夹列表
function file_list(path)
  if not file_exists(path) then file_list=array("","") : exit function
  dim fso : set fso = Server.CreateObject(OBJ_FSO)
  dim fdr : set fdr = fso.getfolder(Server.MapPath(path))
  dim folders : set folders = fdr.subfolders
  dim f, t, tmp : t = "" : tmp = ""
  for each f in folders
    tmp = tmp & t & f.name : t = "|"
  next
  tmp = tmp & "*" : t = ""
  dim files : set files = fdr.files
  for each f in files
    tmp = tmp & t & f.name : t = "|"
  next
  set fso = nothing
  file_list = split(tmp,"*")'返回长度为二的字符数组
end function

'函数：返回图片类型及尺寸
function file_imginfo(path)
  dim tmp : tmp = array("",0,0)
  dim fso : set fso = Server.CreateObject(OBJ_FSO)
  if fso.fileexists(Server.MapPath(path)) then
    dim img : set img = loadpicture(Server.MapPath(path))
    select case img.type
      case 0 : tmp(0) = "none"'类型
      case 1 : tmp(0) = "bitmap"
      case 2 : tmp(0) = "metafile"
      case 3 : tmp(0) = "ico"
      case 4 : tmp(0) = "win32-enhanced metafile"
    end select
    tmp(1) = round(img.width/26.4583)'宽度
    tmp(2) = round(img.height/26.4583)'高度
    set img = nothing
    set fso = nothing
  end if
  file_imginfo = tmp
end function

'函数：检测图片文件合法性
function file_isimg(path)
  dim tmp : tmp = false
  if not file_exists(path) then file_isimg = tmp : exit function
  dim jpg(1):jpg(0)=cbyte(&HFF):jpg(1)=cbyte(&HD8)
  dim bmp(1):bmp(0)=cbyte(&H42):bmp(1)=cbyte(&H4D)
  dim png(3):png(0)=cbyte(&H89):png(1)=cbyte(&H50):png(2)=cbyte(&H4E):png(3)=cbyte(&H47)
  dim gif(5):gif(0)=cbyte(&H47):gif(1)=cbyte(&H49):gif(2)=cbyte(&H46):gif(3)=cbyte(&H39):gif(4)=cbyte(&H38):gif(5)=cbyte(&H61)
  dim fstream,fext,stamp,i
  fext = mid(path, instrrev(path,".") + 1)
  set fstream = Server.CreateObject(OBJ_STRM)
  fstream.open
  fstream.type = 1
  fstream.loadfromfile Server.MapPath(path)
  fstream.position = 0
  select case fext
    case "jpg","jpeg":
      stamp = fstream.read(2)
      for i = 0 to 1
        if ascb(midb(stamp,i + 1,1)) = jpg(i) then tmp = true else tmp = false
      next
    case "gif":
      stamp = fstream.read(6)
      for i = 0 to 5
        if ascb(midb(stamp,i + 1,1)) = gif(i) then tmp = true else tmp = false
      next
    case "png":
      stamp = fstream.read(4)
      for i = 0 to 3
        if ascb(midb(stamp,i + 1,1)) = png(i) then tmp = true else tmp = false
      next
    case "bmp":
      stamp = fstream.read(2)
      for i = 0 to 1
        if ascb(midb(stamp,i + 1,1)) = bmp(i) then tmp = true else tmp = false
      next
  end select
  fstream.close : set fstream = nothing
  file_isimg = tmp
end function

'函数：采集远程文件并保存到本地磁盘
function file_savefromurl(fileurl,savepath,savetype)
  if savetype <> 1 and savetype <> 2 then savetype = 2
  dim xmlhttp : set xmlhttp = Server.CreateObject(OBJ_XHTP)
  with xmlhttp
    .open "get", fileurl, false
    .send()
    dim fl : fl = .responsebody
  end with
  set xmlhttp = nothing
  dim stream : set stream = Server.CreateObject(OBJ_STRM)
  with stream
    .type = savetype
    .open
    .write fl
    .savetofile Server.MapPath(savepath),2
    .cancel()
    .close()
  end with
  set stream = nothing
  file_savefromurl = file_exists(savepath)
end function

'函数：读取文件内容到字符串
function file_read(path)
  dim tmp : tmp = ""
  if left(path,7) = "http://" then '读取远程文件
    dim xmlhttp : set xmlhttp = Server.CreateObject(OBJ_XHTP)
    with xmlhttp
      .open "get", path, false
      .send()
      tmp = .responsetext
    end with
    set xmlhttp = nothing
  else '读取本地文件
    if not file_exists(path) then file_read = tmp : exit function
    dim stream : set stream = Server.CreateObject(OBJ_STRM)
    with stream
      .type = 2 '文本类型
      .mode = 3 '读写模式
      .charset = EASYIDE_CHARSET
      .open
      .loadfromfile(Server.MapPath(path))
      tmp = .readtext()
    end with
    stream.close : set stream = nothing
  end if
  file_read = tmp
end function

'函数：保存字符串到文件
function file_save(str,path,model)
  if model <> 0 and model <> 1 then model = 1
  if model = 0 and file_exists(path) then file_save = true : exit function
  dim stream : set stream = Server.CreateObject(OBJ_STRM)
  with stream
    .type = 2 '文本类型
    .charset = EASYIDE_CHARSET
    .open
    .writetext str
    .savetofile(Server.MapPath(path)),model + 1
  end with
  stream.close : set stream = nothing
  file_save = file_exists(path)
end function

'函数:读取ASP类型文件的全部内容
function file_iread(path)
  dim str : str = file_read(path)
  dim pattern : pattern = "<\!--#include[ ]+?file[ ]*?=[ ]*?""(\S+?)""--\>"
  dim matches : set matches = str_execute(pattern,str)
  dim m, f, tmp
  for each m in matches
    f = mid(path,1,instrrev(path,"/")) & m.submatches(0)
    tmp = file_read(f)
    if str_test(pattern,tmp) then tmp = file_iread(f) '处理子包含
    str = replace(str,m.value,tmp)
  next
  pattern = "<%@[ ]*?LANGUAGE[ ]*?=[ ]*?""[a-zA-Z]+?""[ ]+?CODEPAGE[ ]*?=[ ]*?""[0-9]+?""[ ]*?%\>"
  str = str_replace(pattern,str,"")
  file_iread = str
end function

'函数：将ASP文件运行结果返回为字串
function file_eval(path)
  dim str : str = file_iread(path)
  file_eval = str_eval(str)
end function


'/////////数组操作函数部分

'函数：检测元素是否是指定数组的元素成员
function arr_in(arr,val)
  dim a, tmp : tmp = false
  for each a in arr
    if trim(a) = trim(val) then : tmp = true : exit for
  next
  arr_in = tmp
end function

'函数：指定字串数组的元素是否含有指定字串
function arr_strin(arr,str)
  dim a, tmp : tmp = false
  for each a in arr
    if instr(1,a,str,1) <> 0 then : tmp = true : exit for
  next
  arr_strin = tmp
end function

'函数：动态向数组中添加新元素
function arr_push(arr,val)
  redim preserve arr(ubound(arr) + 1)
  arr(ubound(arr)) = val
  arr_push = arr
end function

'函数：获取元素在数组中首次出现时的索引值
function arr_getindex(arr,str)
  dim i, tmp : tmp = -1
  for i = 0 to ubound(arr)
    if arr(i) = str then tmp = i : exit for
  next
  arr_getindex = tmp
end function


'/////////XML解析操作函数部分

'函数：载入xml文件并返回操作对象
function xml_load(path)
  dim obj_xml
  set obj_xml = Server.CreateObject(OBJ_DOM)
  obj_xml.load Server.MapPath(path)
  set xml_load = obj_xml
end function


'/////////数据操作函数部分

'函数：执行SQL语句
function ado_query(byval sql)
  set ado_query = ado_iquery(sql,conn,3,1)
end function

'函数：执行SQL语句,可修改
function ado_query_modify(byval sql)
  set ado_query_modify = ado_iquery(sql,conn,3,2)
end function

'函数：执行SQL语句
function ado_iquery(byval sql,conn,cursortype,locktype)
  if trim(sql) = "" then exit function
  dim rs
  if lcase(left(ltrim(sql),6)) = "select" then
    set rs = Server.CreateObject(OBJ_RST)
    rs.cursorlocation = 3
    rs.open sql,conn,cursortype,locktype
  else
    set rs = conn.execute(sql)
  end if
  set ado_iquery = rs
end function


'/////////翻页操作函数部分

'函数：翻页预处理
function pageturner_handle(byval sql,field_id,page_size)
  pageturner_handle = pageturner_ihandle(sql,field_id,page_size,conn)
end function

'函数：翻页预处理
function pageturner_ihandle(sql,field_id,page_size,conn)
  '获取总记录数：page_sum
  dim rs, page_sum, page_num
  set rs = ado_iquery(sql,conn,3,1)
  page_sum = rs.recordcount
  '计算总页数：page_num
  rs.pagesize = page_size
  page_num = rs.pagecount
  '获取翻页参数
  dim page : page = Request("page")
  if isempty(page) or not isnumeric(page) then page = 1
  if cdbl(page) <= 0 then page = 1
  if cdbl(page) > cdbl(page_num) then page = page_num
  '获取当前页ID列表
  dim i, s, tmp, filter : s = "" : filter = field_id & "="
  if not rs.eof then rs.absolutepage = page
  for i = 1 to page_size
    tmp = ""
    if not rs.eof then
      if not isnumeric(rs(field_id)) then tmp = "'"
      filter = filter & s & tmp & rs(field_id) & tmp
      s = " or " & field_id & "="
      rs.movenext
    end if
  next
  'die filter
  if page_sum > 0 then rs.filter = filter
  '返回数组
  pageturner_ihandle = array(rs,page,page_num,page_sum)
end function

'函数：返回翻页条
function pageturner_show(page,page_num,page_sum,page_size,page_len)
  dim page_start, page_end, page_link, tmp, p
  '起始页、结束页
  page_start = page - page_len
  page_end = page + page_len
  if cdbl(page_start) <= 0 then
    page_end = page_end + abs(page_start)
    page_start = 1
  end if
  if cdbl(page_end) > cdbl(page_num) then page_end = page_num
  '翻页链接
  'page_link="?" : if str_query("page")<>"" then page_link = "?" & str_query("page") & "&"
  page_link = "?" : tmp = str_query("page"): if tmp<>"" then page_link = "?"&tmp&"&"
  '翻页条开始
  dim page_back, page_next
  tmp = "<div class=""page_turner"">"
  if cdbl(page) = 1 then
    page_back = "<a title=""上一页"" href=""javascript:void(0)"">&#8249;&#8249;</a>"
  else
    page_back = "<a title=""上一页"" href="""& page_link & "page="& (page-1) &""">&#8249;&#8249;</a>"
  end if'上一页
  if cdbl(page) > page_len+1 then tmp = tmp & "<a title=""首页"" href="""& page_link & "page=1"">1...</a>"'首页
  for p = page_start to page_end
    if cdbl(p) = cdbl(page) then
      tmp = tmp & "<a title=""第"& p &"页"" class=""c"">"& p &"</a>"
    else
      tmp = tmp & "<a title=""第"& p &"页"" href="""& page_link &"page="& p &""">"& p &"</a>"
    end if
  next'第_页
  if cdbl(page) = cdbl(page_num) then
    page_next = "<a title=""下一页"" href=""javascript:void(0)"">&#8250;&#8250;</a>"
  else
    page_next = "<a title=""下一页"" href="""& page_link & "page="& (page+1) &""">&#8250;&#8250;</a>"
  end if'下一页
  if cdbl(page)<cdbl(page_num)-page_len then tmp = tmp&"<a title=""末页"" href="""&page_link&"page="& page_num &""">..."&page_num&"</a>"'末页
  tmp = tmp & page_back & page_next
  tmp = tmp & "<span>"& page_size &"条<cite>/</cite>页&nbsp;共<label id=""total"">"& page_sum &"</label>条</span>"
  tmp = tmp & "</div>"
  pageturner_show = tmp
end function


'/////////base6 class for VBs

class base64_class
  private blen_
  private bstr_
  
  public property get bstr
    bstr = bstr_
  end property
  
  public property let bstr(val)
    bstr_ = val
  end property
  
  public property get blen
    blen = blen_
  end property
  
  public property let blen(val)
    blen_ = val
  end property
  
  private sub class_initialize
    bstr = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    blen = 8
  end sub
  
  'private sub class_terminate
  'end sub
    
  public function cbit(num)
    dim cbitstr : cbitstr = ""
    if len(num)>0 and isnumeric(num) then
      do while not num\2 < 1
        cbitstr = (num mod 2) & cbitstr
        num = num\2
      loop
    end if
    cbit = num & cbitstr
  end function
  
  public function cdec(num)
    dim inum, cdecstr : cdecstr = 0
    if len(num)>0 and isnumeric(num) then
      for inum=0 to len(num)-1
        cdecstr = cdecstr + 2^inum*cint(mid(num,len(num)-inum,1))
      next
    end if
    cdec = cdecstr
  end function
  
  public function encode(str)
    if not len(str)>0 then exit function
    dim i, t, s, encodestr
    t = ""
    s = ""
    encodestr = ""
    for i=1 to len(str)
      't = abs(ascw(mid(str,i,1)))
      t = ascw(mid(str,i,1))
      if t<0 then t = t + 65536
      t = cbit(t)
      if len(t)<blen then t = string(blen-len(t),"0") & t
      s = s & t
    next
    if len(s) mod 6 <> 0 then s = s & string(6-(len(s) mod 6),"0")
    t = ""
    for i=1 to len(s)\6
      t = cdec(mid(s,i*6-6+1,6))
      encodestr = encodestr & mid(bstr,t+1,1)
    next
    if len(encodestr)<4 then encodestr = encodestr & string(4-len(encodestr),"=")
    encode = encodestr
  end function
  
  public function decode(str)
    if not len(str)>0 then exit function
    dim i, t, s, decodestr
    t = ""
    s = ""
    decodestr = ""
    str = replace(str,"=","")
    for i=1 to len(str)
      t = cbit(instr(bstr,mid(str,i,1)) - 1)
      if len(t)<6 then t = string(6-len(t),"0") & t
      s = s & t
    next
    if len(s) mod blen <> 0 then  s = left(s,len(s)-(len(s) mod blen))
    t = ""
    for i = 1 to len(s)\blen
      t = cdec(mid(s,i*blen-blen+1,blen))
      decodestr = decodestr & chrw(t)
    next
    decode = decodestr
  end function
end class
%>