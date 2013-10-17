<%
'=====================================================================================
' Document :  EasyIDE Library [EasyIDE ASP 函数库]
' Version  :  V2.1
' Author   :  沉沦
' Contact  :  www.EasyIDE.cn
' Update   :  2009-05-12
'=====================================================================================

const OBJ_RST  = "ADODB.Recordset"
const OBJ_CONN = "ADODB.Connection"
const OBJ_STRM = "ADODB.Stream"
const OBJ_FSO  = "Scripting.FilesyStemObject"
const OBJ_XHTP = "MSXML2.XMLHTTP"
const OBJ_DOM  = "MSXML2.DOMDocument"

'/////////基础操作函数部分

'过程：输出字符串[代替Response.Write]
sub echo(str)
    response.Write(str)
end sub

'过程：结束页面并输出字符串
sub die(str)
    response.Write(str) : response.End()
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
    if not isnumeric(num) then ctime=tmp : exit function
    if n="" or not isnumeric(n) then n = 2
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
    if not isnumeric(num) then csize=tmp : exit function
    if n="" or not isnumeric(n) then n = 2
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

'函数：将ASP文件运行结果返回为字串
function ob_get_contents(path)
    dim tmp, a, b, t, matches, m
    dim str : str = file_iread(path)
    tmp = "dim htm : htm = """""&vbcrlf
    a = 1
    b = instr(a,str,"<%")+2
    while b > a+1
        t = mid(str,a,b-a-2)
        t = replace(t,vbcrlf,"{::vbcrlf}")
        t = replace(t,vbcr,"{::vbcr}")
        t = replace(t,"""","""""")
        tmp = tmp & "htm = htm & """ & t & """" & vbcrlf
        a = instr(b,str,"%\>")+2
        tmp = tmp & str_replace("^\s*=",mid(str,b,a-b-2),"htm = htm & ") & vbcrlf
        b = instr(a,str,"<%")+2
    wend
    t = mid(str,a)
    t = replace(t,vbcrlf,"{::vbcrlf}")
    t = replace(t,vbcr,"{::vbcr}")
    t = replace(t,"""","""""")
    tmp = tmp & "htm = htm & """ & t & """" & vbcrlf
    tmp = replace(tmp,"response.write","htm = htm & ",1,-1,1)
    tmp = replace(tmp,"echo","htm = htm & ",1,-1,1)
    'execute(tmp)
    executeglobal(tmp)
    htm = replace(htm,"{::vbcrlf}",vbcrlf)
    htm = replace(htm,"{::vbcr}",vbcr)
    ob_get_contents = htm
end function

'过程：动态包含文件
sub include(path)
    echo ob_get_contents(path)
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
    str = server.URLEncode(str)
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

'函数：md5加密
function md5(byval str)
    if isnull(str) then exit function
    dim md5_cls
    set md5_cls = new md5_class
    str = md5_cls.md5(str)
    set md5_cls = nothing
    md5 = str
end function

'函数：三元IF
function iif(exp,v1,v2)
    dim tmp : tmp = v2
    if exp then tmp = v1
    iif = tmp
end function

'函数：空值测试
function inull(val)
    dim tmp : tmp = false
    if isnull(val) then
        tmp = true
    elseif isempty(val) then
        tmp = true
    elseif trim(val)="" then
        tmp = true
    end if
    inull = tmp
end function

'全启变量：客户端IP
dim ip : ip = request.ServerVariables("REMOTE_ADDR")


'函数：返回客户端真实IP
function realip
    dim tmp : tmp = request.ServerVariables("HTTP_X_FORWARDED_FOR")
    if trim(tmp)="" then tmp = request.ServerVariables("REMOTE_ADDR")
    realip = tmp
end function

'函数：邮件发送[Jamil-Message]
function sendmail(fromname,sendto,subject,body,from,serveraddress,username,password)
    dim jmail, return
    set jmail        = server.CreateObject("JMAIL.Message")
    jmail.silent        = true
    jmail.logging        = true
    jmail.charset        = "utf-8"
    jmail.contenttype    = "text/html; charset=utf-8"
    jmail.addrecipient sendto
    jmail.fromname        = fromname
    jmail.from        = from
    jmail.mailserverusername= username
    jmail.MailServerPassword= password
    jmail.subject        = subject
    jmail.body        = body
    jmail.priority        = 3
    return            = jmail.send(serveraddress)
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
    set obj_test = server.CreateObject(str)
    if err.number = 0 then tmp = true
    set obj_test = nothing
    err.clear()
    install = tmp
end function

'/////////字符串操作函数部分

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
    dim tmp : tmp = "&quot;=""|&amp;=&|&lt;=<|&gt;=>|&euro;=€|&nbsp;= |&laquo;=«|&raquo;=»|&hellip;=…|&copy;=©"
    dim arr, a, v : arr = split(tmp,"|")
    for each v in arr
        a = split(v,"=")
        str = replace(str,a(0),a(1))
    next
    'die str
    dim i, c, s, n : n = 0 : tmp = ""
    for i=1 to len(str)
        s = mid(str,i,1)
        c = abs(ascw(s))
        if c>255 then n=n+2 else n=n+1
        tmp = tmp & s
        if n >= slen then exit for
    next
    if tmp=str then ext=""
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
    dim pattern : pattern = "<[^>]+?>"
    str = str_replace(pattern,str,"")
    str_htmlclear = str
end function

'函数：清除所有格式及空格 压缩字符串
function str_trim(byval str)
    if isnull(str) then exit function
    str = replace(str,chr(10),"")
    str = replace(str,chr(13),"")
    dim pattern : pattern = "<[^>]+?>"
    str = str_replace(pattern,str,"")
    str = replace(str,"&nbsp;","")
    str = replace(str," ","")
    str_trim = str
end function

'函数：返回一个不重复的随机字串
function str_rnd()
    dim ran_num, dt_now, tmp : dt_now = now()
    randomize : ran_num = int( (90000*rnd) + 10000 )
    tmp = year(dt_now) & right("0"&month(dt_now),2) & right("0"&day(dt_now),2) & right("0"&hour(dt_now),2) &_
          right("0"&minute(dt_now),2) & right("0"&second(dt_now),2) & ran_num
    str_rnd = base64encode(tmp)
end function

'函数：返回格式化的时间字串
function str_time(format,byval str)
    if trim(str)="" or not isdate(str) then exit function
    dim tmp : tmp = format
    tmp = replace(tmp,"yy",right("0"&year(str),2),1,-1,1)
    tmp = replace(tmp,"y",year(str),1,-1,1)
    tmp = replace(tmp,"mm",right("0"&month(str),2),1,-1,1)
    tmp = replace(tmp,"m",month(str),1,-1,1)
    tmp = replace(tmp,"dd",right("0"&day(str),2),1,-1,1)
    tmp = replace(tmp,"d",day(str),1,-1,1)
    tmp = replace(tmp,"hh",right("0"&hour(str),2),1,-1,1)
    tmp = replace(tmp,"h",hour(str),1,-1,1)
    tmp = replace(tmp,"ii",right("0"&minute(str),2),1,-1,1)
    tmp = replace(tmp,"i",minute(str),1,-1,1)
    tmp = replace(tmp,"ss",right("0"&second(str),2),1,-1,1)
    tmp = replace(tmp,"s",second(str),1,-1,1)
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
    dim tmp : tmp = request.ServerVariables("QUERY_STRING")
    if trim(del) = "" then str_query = tmp : exit function
    dim arr : arr = split(tmp, "&")
    dim q, a, t : t = "" : tmp = ""
    for each q in arr
        if trim(q) <> "" then
            a = split(q,"=") : if ubound(a)=0 then arr_push a,""
            if not arr_in(split(del,","),a(0)) then tmp = tmp&t&a(0)&"="&a(1) : t = "&"
        end if
    next
    str_query = tmp
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

'/////////文件操作函数部分

'函数：获取当前脚本执行文件的文件名
function file_self()
    dim tmp
    tmp = request.ServerVariables("SCRIPT_NAME")
    tmp = split(tmp,"/")
    file_self = tmp(ubound(tmp))
end function

'函数：获取当前脚本执行文件所在的磁盘目录
function file_dir()
    dim tmp, arr
    tmp = request.ServerVariables("SCRIPT_NAME")
    arr = split(tmp,"/")
    tmp = arr(ubound(arr))
    arr = split(server.MapPath(tmp),"\")
    file_dir = arr(ubound(arr)-1)
end function

'函数：检测文件/文件夹是否存在
function file_exists(path)
    dim tmp : tmp = false
    dim fso : set fso = server.CreateObject(OBJ_FSO)
    if fso.fileexists(server.MapPath(path)) then tmp = true
    if fso.folderexists(server.MapPath(path)) then tmp = true
    set fso = nothing
    file_exists = tmp
end function

'函数：删除文件/文件夹
function file_delete(path)
    dim tmp : tmp = false
    dim fso : set fso=server.CreateObject(OBJ_FSO)
    if fso.fileexists(server.MapPath(path)) then'目标是文件
        fso.deletefile(server.MapPath(path))
        if not fso.fileexists(server.MapPath(path)) then tmp = true
    end if
    if fso.folderexists(server.MapPath(path)) then'目标是文件夹
        fso.deletefolder(server.MapPath(path))
        if not fso.folderexists(server.MapPath(path)) then tmp = true
    end if
    set fso = nothing
    file_delete = tmp
end function

'函数：获取文件/文件夹信息
function file_info(path)
    dim tmp(4)
    dim fso : set fso = server.CreateObject(OBJ_FSO)
    if fso.fileexists(server.MapPath(path)) then '目标是文件
        dim fl : set fl = fso.getfile(server.MapPath(path))
        tmp(0) = fl.type'类型
        tmp(1) = fl.attributes'属性
        tmp(2) = csize(fl.size,4)'大小
        tmp(3) = fl.datecreated'创建时间
        tmp(4) = fl.datelastmodified'最后修改时间
    elseif fso.folderexists(server.MapPath(path)) then '目标是文件夹
        dim fd : set fd = fso.getfolder(server.MapPath(path))
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
    if model<>0 and model<>1 then model = false else model = cbool(model)
    dim tmp : tmp = false
    dim fso : set fso = server.CreateObject(OBJ_FSO)
    if fso.fileexists(server.MapPath(file_start)) then '目标是文件
        fso.copyfile server.MapPath(file_start),server.MapPath(file_end),model
        if fso.fileexists(server.MapPath(file_end)) then tmp = true
    end if
    if fso.folderexists(server.MapPath(file_start)) then '目标是文件夹
        fso.copyfolder server.MapPath(file_start),server.MapPath(file_end),model
        if fso.folderexists(server.MapPath(file_end)) then tmp = true
    end if
    set fso = nothing
    file_copy = tmp
end function

'函数：创建文件夹
function file_create(path,model)
    if model<>0 and model<>1 then model = false else model = cbool(model)
    dim tmp : tmp = false
    dim fso : set fso = server.CreateObject(OBJ_FSO)
    if fso.folderexists(server.MapPath(path)) then
        if model then fso.deletefolder(server.MapPath(path)) : fso.createfolder server.MapPath(path)
    else
        fso.createfolder server.MapPath(path)
    end if
    if fso.folderexists(server.MapPath(path)) then tmp = true
    set fso = nothing
    file_create = tmp
end function

'函数：获取指定目录下所有文件及文件夹列表
function file_list(path)
    if not file_exists(path) then file_list=array("","") : exit function
    dim fso : set fso = server.CreateObject(OBJ_FSO)
    dim fdr : set fdr = fso.getfolder( server.MapPath(path) )
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
    dim fso : set fso = server.CreateObject(OBJ_FSO)
    if fso.fileexists(server.MapPath(path)) then
        dim img : set img = loadpicture(server.MapPath(path))
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
    fext = mid(path, instrrev(path,".")+1)
    set fstream = server.CreateObject(OBJ_STRM)
    fstream.open
    fstream.type = 1
    fstream.loadfromfile server.MapPath(path)
    fstream.position = 0
    select case fext
        case "jpg","jpeg":
            stamp = fstream.read(2)
            for i=0 to 1
                if ascb(midb(stamp,i+1,1))=jpg(i) then tmp=true else tmp=false
            next
        case "gif":
            stamp = fstream.read(6)
            for i=0 to 5
                if ascb(midb(stamp,i+1,1))=gif(i) then tmp=true else tmp=false
            next
        case "png":
            stamp = fstream.read(4)
            for i=0 to 3
                if ascb(midb(stamp,i+1,1))=png(i) then tmp=true else tmp=false
            next
        case "bmp":
            stamp = fstream.read(2)
            for i=0 to 1
                if ascb(midb(stamp,i+1,1))=bmp(i) then tmp=true else tmp=false
            next
    end select
    fstream.close : set fstream = nothing
    file_isimg = tmp
end function

'函数：采集远程文件并保存到本地磁盘
function file_savefromurl(fileurl,savepath,savetype)
    if savetype<>1 and savetype<>2 then savetype=2
    dim xmlhttp : set xmlhttp = server.CreateObject(OBJ_XHTP)
    with xmlhttp
        .open "get", fileurl, false
        .send()
        dim fl : fl = .responsebody
    end with
    set xmlhttp = nothing
    dim stream : set stream = server.CreateObject(OBJ_STRM)
    with stream
        .type = savetype
        .open
        .write fl
        .savetofile server.MapPath(savepath), 2
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
        dim xmlhttp : set xmlhttp = server.CreateObject(OBJ_XHTP)
        with xmlhttp
            .open "get", path, false
            .send()
            tmp = .responsetext
        end with
        set xmlhttp = nothing
    else '读取本地文件
        if not file_exists(path) then file_read = tmp : exit function
        dim stream : set stream = server.CreateObject(OBJ_STRM)
        with stream
            .type = 2 '文本类型
            .mode = 3 '读写模式
            .charset = "utf-8"
            .open
            .loadfromfile(server.MapPath(path))
            tmp = .readtext()
        end with
        stream.close : set stream = nothing
    end if
    file_read = tmp
end function

'函数：保存字符串到文件
function file_save(str,path,model)
    if model<>0 and model<>1 then model=1
    if model=0 and file_exists(path) then file_save=true : exit function
    dim stream : set stream = server.CreateObject(OBJ_STRM)
    with stream
        .type = 2 '文本类型
        .charset = "utf-8"
        .open
        .writetext str
        .savetofile(server.MapPath(path)),model+1
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
        f = mid(path,1,instrrev(path,"/"))&m.submatches(0)
        tmp = file_read(f)
        if str_test(pattern,tmp) then tmp = file_iread(f) '处理子包含
        str = replace(str,m.value,tmp)
    next
    pattern = "<%@[ ]*?LANGUAGE[ ]*?=[ ]*?""[a-zA-Z]+?""[ ]+?CODEPAGE[ ]*?=[ ]*?""[0-9]+?""[ ]*?%\>"
    str = str_replace(pattern,str,"")
    file_iread = str
end function

'/////////数组操作函数部分

'函数：检测元素是否是指定数组的元素成员
function arr_in(arr,val)
    dim a, tmp : tmp = false
    for each a in arr
        if trim(a)=trim(val) then : tmp=true : exit for
    next
    arr_in = tmp
end function

'函数：指定字串数组的元素是否含有指定字串
function arr_strin(arr,str)
    dim a, tmp : tmp = false
    for each a in arr
        if instr(1,a,str,1)<>0 then : tmp=true : exit for
    next
    arr_strin = tmp
end function

'函数：动态向数组中添加新元素
function arr_push(arr,val)
    redim preserve arr(ubound(arr)+1)
    arr(ubound(arr)) = val
    arr_push = arr
end function

'函数：获取元素在数组中首次出现时的索引值
function arr_getindex(arr,str)
    dim i, tmp : tmp = -1
    for i=0 to ubound(arr)
        if arr(i) = str then tmp = i : exit for
    next
    arr_getindex = tmp
end function

'/////////XML解析操作函数部分

'函数：载入xml文件并返回操作对象
function xml_load(path)
    dim obj_xml
    set obj_xml = server.CreateObject(OBJ_DOM)
    obj_xml.load Server.MapPath(path)
    set xml_load = obj_xml
end function

'/////////数据操作函数部分

'函数：执行SQL语句
function ado_query(byval sql)
    set ado_query = ado_iquery(sql,conn,3,1)
end function

'函数：执行SQL语句
function ado_iquery(byval sql,conn,cursortype,locktype)
    if trim(sql) = "" then exit function
    if trim(n)="" or not isnumeric(n) then n=1
    dim rs
    if lcase(left(ltrim(sql),6)) = "select" then
        set rs = server.CreateObject(OBJ_RST)
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
    pageturner_handle =pageturner_ihandle(sql,field_id,page_size,conn)
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
    dim page : page = request.QueryString("page")
    if isempty(page) or not isnumeric(page) then page = 1
    if cdbl(page) <= 0 then page = 1
    if cdbl(page) > cdbl(page_num) then page = page_num
    '获取当前页ID列表
    dim i, s, filter : s = "" : filter = field_id&"="
    if not rs.eof then rs.absolutepage = page
    for i = 1 to page_size
        if not rs.eof then
            filter = filter & s & rs(field_id)
            s = " or "&field_id&"="
            rs.movenext
        end if
    next
    'die filter
    if page_sum>0 then rs.filter = filter
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

'/////////md5 class for VBs

class md5_class
    private BITS_TO_A_BYTE
    private BYTES_TO_A_WORD
    private BITS_TO_A_WORD
    private m_lOnBits(30)
    private m_l2Power(30)
    
    private sub class_initialize
        BITS_TO_A_BYTE = 8
        BYTES_TO_A_WORD = 4
        BITS_TO_A_WORD = 32
    end sub
    
    private function LShift(lValue, iShiftBits)
        if iShiftBits = 0 then
            LShift = lValue
            exit function
        elseif iShiftBits = 31 then
            if lValue and 1 then
                LShift = &H80000000
            else
                LShift = 0
            end If
            exit function
        elseif iShiftBits < 0 or iShiftBits > 31 then
            err.raise 6
        end If
        if (lValue and m_l2Power(31 - iShiftBits)) then
            LShift = ((lValue and m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) or &H80000000
        else
            LShift = ((lValue and m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
        end If
    end function
    
    private function RShift(lValue, iShiftBits)
        if iShiftBits = 0 then
            RShift = lValue
            exit function
        elseif iShiftBits = 31 then
            if lValue and &H80000000 then
                RShift = 1
            else
                RShift = 0
            end If
            exit function
        elseif iShiftBits < 0 or iShiftBits > 31 then
            err.raise 6
        end If
        RShift = (lValue and &H7FFFFFFE) \ m_l2Power(iShiftBits)
        if (lValue and &H80000000) then
            RShift = (RShift or (&H40000000 \ m_l2Power(iShiftBits - 1)))
        end If
    end function
    
    private function RotateLeft(lValue, iShiftBits)
        RotateLeft = LShift(lValue, iShiftBits) or RShift(lValue, (32 - iShiftBits))
    end function
    
    private function AddUnsigned(lX, lY)
        dim lX4
        dim lY4
        dim lX8
        dim lY8
        dim lResult
        
        lX8 = lX and &H80000000
        lY8 = lY and &H80000000
        lX4 = lX and &H40000000
        lY4 = lY and &H40000000
        
        lResult = (lX and &H3FFFFFFF) + (lY and &H3FFFFFFF)
        
        if lX4 and lY4 then
            lResult = lResult xor &H80000000 xor lX8 xor lY8
        elseif lX4 or lY4 then
            if lResult and &H40000000 then
                lResult = lResult xor &HC0000000 xor lX8 xor lY8
            else
                lResult = lResult xor &H40000000 xor lX8 xor lY8
            end If
        else
            lResult = lResult xor lX8 xor lY8
        end If
        AddUnsigned = lResult
    end function
    
    private function md5_F(x, y, z)
        md5_F = (x and y) or ((not x) and z)
    end function
    
    private function md5_G(x, y, z)
        md5_G = (x and z) or (y and (not z))
    end function
    
    private function md5_H(x, y, z)
        md5_H = (x xor y xor z)
    end function
    
    private function md5_I(x, y, z)
        md5_I = (y xor (x or (not z)))
    end function
    
    private sub md5_FF(a, b, c, d, x, s, ac)
        a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_F(b, c, d), x), ac))
        a = RotateLeft(a, s)
        a = AddUnsigned(a, b)
    end sub
    
    private sub md5_GG(a, b, c, d, x, s, ac)
        a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_G(b, c, d), x), ac))
        a = RotateLeft(a, s)
        a = AddUnsigned(a, b)
    end sub
    
    private sub md5_HH(a, b, c, d, x, s, ac)
        a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_H(b, c, d), x), ac))
        a = RotateLeft(a, s)
        a = AddUnsigned(a, b)
    end sub
    
    private sub md5_II(a, b, c, d, x, s, ac)
        a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_I(b, c, d), x), ac))
        a = RotateLeft(a, s)
        a = AddUnsigned(a, b)
    end sub
    
    private function ConvertToWordArray(sMessage)
        dim lMessageLength
        dim lNumberOfWords
        dim lWordArray()
        dim lBytePosition
        dim lByteCount
        dim lWordCount
        dim MODULUS_BITS : MODULUS_BITS = 512
        dim CONGRUENT_BITS : CONGRUENT_BITS = 448
        lMessageLength = Len(sMessage)
        lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
        redim lWordArray(lNumberOfWords - 1)
        lBytePosition = 0
        lByteCount = 0
        do until lByteCount >= lMessageLength
            lWordCount = lByteCount \ BYTES_TO_A_WORD
            lBytePosition = (lByteCount mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
            lWordArray(lWordCount) = lWordArray(lWordCount) or LShift(Asc(Mid(sMessage, lByteCount + 1, 1)), lBytePosition)
            lByteCount = lByteCount + 1
        loop
        lWordCount = lByteCount \ BYTES_TO_A_WORD
        lBytePosition = (lByteCount mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
        lWordArray(lWordCount) = lWordArray(lWordCount) or LShift(&H80, lBytePosition)
        lWordArray(lNumberOfWords - 2) = LShift(lMessageLength, 3)
        lWordArray(lNumberOfWords - 1) = RShift(lMessageLength, 29)
        ConvertToWordArray = lWordArray
    end function
    
    private function WordToHex(lValue)
        dim lByte
        dim lCount
        for lCount = 0 to 3
            lByte = RShift(lValue, lCount * BITS_TO_A_BYTE) and m_lOnBits(BITS_TO_A_BYTE - 1)
            WordToHex = WordToHex & Right("0" & Hex(lByte), 2)
        next
    end function
    
    public function MD5(sMessage)
        m_lOnBits(0) = CLng(1)
        m_lOnBits(1) = CLng(3)
        m_lOnBits(2) = CLng(7)
        m_lOnBits(3) = CLng(15)
        m_lOnBits(4) = CLng(31)
        m_lOnBits(5) = CLng(63)
        m_lOnBits(6) = CLng(127)
        m_lOnBits(7) = CLng(255)
        m_lOnBits(8) = CLng(511)
        m_lOnBits(9) = CLng(1023)
        m_lOnBits(10) = CLng(2047)
        m_lOnBits(11) = CLng(4095)
        m_lOnBits(12) = CLng(8191)
        m_lOnBits(13) = CLng(16383)
        m_lOnBits(14) = CLng(32767)
        m_lOnBits(15) = CLng(65535)
        m_lOnBits(16) = CLng(131071)
        m_lOnBits(17) = CLng(262143)
        m_lOnBits(18) = CLng(524287)
        m_lOnBits(19) = CLng(1048575)
        m_lOnBits(20) = CLng(2097151)
        m_lOnBits(21) = CLng(4194303)
        m_lOnBits(22) = CLng(8388607)
        m_lOnBits(23) = CLng(16777215)
        m_lOnBits(24) = CLng(33554431)
        m_lOnBits(25) = CLng(67108863)
        m_lOnBits(26) = CLng(134217727)
        m_lOnBits(27) = CLng(268435455)
        m_lOnBits(28) = CLng(536870911)
        m_lOnBits(29) = CLng(1073741823)
        m_lOnBits(30) = CLng(2147483647)
        m_l2Power(0) = CLng(1)
        m_l2Power(1) = CLng(2)
        m_l2Power(2) = CLng(4)
        m_l2Power(3) = CLng(8)
        m_l2Power(4) = CLng(16)
        m_l2Power(5) = CLng(32)
        m_l2Power(6) = CLng(64)
        m_l2Power(7) = CLng(128)
        m_l2Power(8) = CLng(256)
        m_l2Power(9) = CLng(512)
        m_l2Power(10) = CLng(1024)
        m_l2Power(11) = CLng(2048)
        m_l2Power(12) = CLng(4096)
        m_l2Power(13) = CLng(8192)
        m_l2Power(14) = CLng(16384)
        m_l2Power(15) = CLng(32768)
        m_l2Power(16) = CLng(65536)
        m_l2Power(17) = CLng(131072)
        m_l2Power(18) = CLng(262144)
        m_l2Power(19) = CLng(524288)
        m_l2Power(20) = CLng(1048576)
        m_l2Power(21) = CLng(2097152)
        m_l2Power(22) = CLng(4194304)
        m_l2Power(23) = CLng(8388608)
        m_l2Power(24) = CLng(16777216)
        m_l2Power(25) = CLng(33554432)
        m_l2Power(26) = CLng(67108864)
        m_l2Power(27) = CLng(134217728)
        m_l2Power(28) = CLng(268435456)
        m_l2Power(29) = CLng(536870912)
        m_l2Power(30) = CLng(1073741824)
        dim x
        dim k
        dim AA
        dim BB
        dim CC
        dim DD
        dim a
        dim b
        dim c
        dim d
        dim S11 : S11 = 7
        dim S12 : S12 = 12
        dim S13 : S13 = 17
        dim S14 : S14 = 22
        dim S21 : S21 = 5
        dim S22 : S22 = 9
        dim S23 : S23 = 14
        dim S24 : S24 = 20
        dim S31 : S31 = 4
        dim S32 : S32 = 11
        dim S33 : S33 = 16
        dim S34 : S34 = 23
        dim S41 : S41 = 6
        dim S42 : S42 = 10
        dim S43 : S43 = 15
        dim S44 : S44 = 21
        x = ConvertToWordArray(sMessage)
        a = &H67452301
        b = &HEFCDAB89
        c = &H98BADCFE
        d = &H10325476
        for k = 0 to UBound(x) Step 16
            AA = a
            BB = b
            CC = c
            DD = d
            md5_FF a, b, c, d, x(k + 0), S11, &HD76AA478
            md5_FF d, a, b, c, x(k + 1), S12, &HE8C7B756
            md5_FF c, d, a, b, x(k + 2), S13, &H242070DB
            md5_FF b, c, d, a, x(k + 3), S14, &HC1BDCEEE
            md5_FF a, b, c, d, x(k + 4), S11, &HF57C0FAF
            md5_FF d, a, b, c, x(k + 5), S12, &H4787C62A
            md5_FF c, d, a, b, x(k + 6), S13, &HA8304613
            md5_FF b, c, d, a, x(k + 7), S14, &HFD469501
            md5_FF a, b, c, d, x(k + 8), S11, &H698098D8
            md5_FF d, a, b, c, x(k + 9), S12, &H8B44F7AF
            md5_FF c, d, a, b, x(k + 10), S13, &HFFFF5BB1
            md5_FF b, c, d, a, x(k + 11), S14, &H895CD7BE
            md5_FF a, b, c, d, x(k + 12), S11, &H6B901122
            md5_FF d, a, b, c, x(k + 13), S12, &HFD987193
            md5_FF c, d, a, b, x(k + 14), S13, &HA679438E
            md5_FF b, c, d, a, x(k + 15), S14, &H49B40821
            md5_GG a, b, c, d, x(k + 1), S21, &HF61E2562
            md5_GG d, a, b, c, x(k + 6), S22, &HC040B340
            md5_GG c, d, a, b, x(k + 11), S23, &H265E5A51
            md5_GG b, c, d, a, x(k + 0), S24, &HE9B6C7AA
            md5_GG a, b, c, d, x(k + 5), S21, &HD62F105D
            md5_GG d, a, b, c, x(k + 10), S22, &H2441453
            md5_GG c, d, a, b, x(k + 15), S23, &HD8A1E681
            md5_GG b, c, d, a, x(k + 4), S24, &HE7D3FBC8
            md5_GG a, b, c, d, x(k + 9), S21, &H21E1CDE6
            md5_GG d, a, b, c, x(k + 14), S22, &HC33707D6
            md5_GG c, d, a, b, x(k + 3), S23, &HF4D50D87
            md5_GG b, c, d, a, x(k + 8), S24, &H455A14ED
            md5_GG a, b, c, d, x(k + 13), S21, &HA9E3E905
            md5_GG d, a, b, c, x(k + 2), S22, &HFCEFA3F8
            md5_GG c, d, a, b, x(k + 7), S23, &H676F02D9
            md5_GG b, c, d, a, x(k + 12), S24, &H8D2A4C8A
            md5_HH a, b, c, d, x(k + 5), S31, &HFFFA3942
            md5_HH d, a, b, c, x(k + 8), S32, &H8771F681
            md5_HH c, d, a, b, x(k + 11), S33, &H6D9D6122
            md5_HH b, c, d, a, x(k + 14), S34, &HFDE5380C
            md5_HH a, b, c, d, x(k + 1), S31, &HA4BEEA44
            md5_HH d, a, b, c, x(k + 4), S32, &H4BDECFA9
            md5_HH c, d, a, b, x(k + 7), S33, &HF6BB4B60
            md5_HH b, c, d, a, x(k + 10), S34, &HBEBFBC70
            md5_HH a, b, c, d, x(k + 13), S31, &H289B7EC6
            md5_HH d, a, b, c, x(k + 0), S32, &HEAA127FA
            md5_HH c, d, a, b, x(k + 3), S33, &HD4EF3085
            md5_HH b, c, d, a, x(k + 6), S34, &H4881D05
            md5_HH a, b, c, d, x(k + 9), S31, &HD9D4D039
            md5_HH d, a, b, c, x(k + 12), S32, &HE6DB99E5
            md5_HH c, d, a, b, x(k + 15), S33, &H1FA27CF8
            md5_HH b, c, d, a, x(k + 2), S34, &HC4AC5665
            md5_II a, b, c, d, x(k + 0), S41, &HF4292244
            md5_II d, a, b, c, x(k + 7), S42, &H432AFF97
            md5_II c, d, a, b, x(k + 14), S43, &HAB9423A7
            md5_II b, c, d, a, x(k + 5), S44, &HFC93A039
            md5_II a, b, c, d, x(k + 12), S41, &H655B59C3
            md5_II d, a, b, c, x(k + 3), S42, &H8F0CCC92
            md5_II c, d, a, b, x(k + 10), S43, &HFFEFF47D
            md5_II b, c, d, a, x(k + 1), S44, &H85845DD1
            md5_II a, b, c, d, x(k + 8), S41, &H6FA87E4F
            md5_II d, a, b, c, x(k + 15), S42, &HFE2CE6E0
            md5_II c, d, a, b, x(k + 6), S43, &HA3014314
            md5_II b, c, d, a, x(k + 13), S44, &H4E0811A1
            md5_II a, b, c, d, x(k + 4), S41, &HF7537E82
            md5_II d, a, b, c, x(k + 11), S42, &HBD3AF235
            md5_II c, d, a, b, x(k + 2), S43, &H2AD7D2BB
            md5_II b, c, d, a, x(k + 9), S44, &HEB86D391
            a = AddUnsigned(a, AA)
            b = AddUnsigned(b, BB)
            c = AddUnsigned(c, CC)
            d = AddUnsigned(d, DD)
        next
        'MD5=LCase(WordToHex(b) & WordToHex(c))
        MD5 = UCase(WordToHex(a) & WordToHex(b) & WordToHex(c) & WordToHex(d))
    end function
end class
%>