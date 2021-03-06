<pre>

此方法可以让Asp数组像PHP数组一样拥有键。
此方法依赖EasyIDE ASP Framework。


方法:array2(arr,k,v)
作用:创建/新增/修改数组元素
参数:arr      数组名
参数:k        键名key
参数:v        值value
返回:array2   array2类型的数组
用法:
  array2 arr,"key","hello would"  '创建一个名为arr的数组，并且有一个键为"key"、值为"hello would"的元素
  array2 arr,0,"hi Alone"         '修改数组中key为0的元素的值为"hi Alone"
  array2 arr,null,"some string"   '向数组arr中添加一个元素，值"some string"，当键为空/空字符串时，key自动为数组的所有key中最大一个数值+1，否则为0，此处key为0


方法:array2_new(arr)
作用:初始化一个空的array2数组
参数:arr      数组名
返回:array2   array2类型的空数组
用法:
  array2_new arr                  '初始化一个名为arr的空数组


方法:new_array2(arr)
作用:同array2_new()，初始化一个空的array2数组


方法:array2_read(arr,k)
作用:根据key读取数组中对应的值
参数:arr      array2数组
参数:k        键名key
返回:         数组中key所对应的值
用法:
  val = array2_read(arr,"key")    '读取数组arr中键为"key"的值
  val = arr("key")                '快捷用法


方法:array2_key(arr)
作用:返回array2数组的所有键
参数:arr      array2数组
返回:array    array2数组的所有键


方法:array2_val(arr)
作用:返回array2数组的所有值
参数:arr      array2数组
返回:array    array2数组的所有值


方法:array2_key_val(arr,idx,key,val)
作用:返回array2数组某个所有的键和值
参数:arr      array2数组
参数:idx      索引
参数:key      存放这个索引的键
参数:val      存放这个索引的值
返回:string   array2数组中这个索引的键
用法:
  array2_key_val arr,0,k,v        '获取数组arr中索引为0(第一个)的键值，并分别存放到k、v两个变量中


方法:array2_ubound(arr)
作用:返回array2数组的下标
参数:arr      array2数组
返回:numeric  数组下标
用法:
  for i = 0 to array2_ubound(arr)
    array2_key_val arr,i,k,v
    Response.Write k & " => " & v &vbnewline
  next


方法:array2_is(arr)
作用:判断是不是array2数组
参数:arr      array2数组
返回:boole    true或false


方法:array2_rs(arr,rs)
作用:将ADODB.Recordset集合转换成array2
参数:arr      数组名
参数:rs       ADODB.Recordset集合
返回:array2   array2类型的数组(二维)
用法:
  ...                             'set rs = ...
  if not rs.eof then
    array2_rs arr,rs              '将Recordset集合转换成名为arr的array2数组
  end if
  ...


方法:array2_rs_row(arr,rs)
作用:将ADODB.Recordset集合的第一行转换成array2
参数:arr      数组名
参数:rs       ADODB.Recordset集合
返回:array2   array2类型的数组


方法:array2_match(str,pat,arr)
作用:执行一个正则表达式匹配
参数:str      要搜索的字符串
参数:pat      正则表达式
参数:arr      数组名
返回:array2   array2类型的数组
用法:
  str = "a=1&b=2"
  pat = "\w+=(\w+)"
  array2_match str,pat,arr
    '返回的为多维数组，结构大致如下：
    arr => (
      0 => (
        0 => a=1,
        1 => 1,
      ),
      1 => (
        0 => b=2,
        1 => 2,
      ),
    )


方法:array2_match_php(str,pat,arr)
作用:执行一个正则表达式匹配(PHP风格)
参数:str      要搜索的字符串
参数:pat      正则表达式
参数:arr      数组名
返回:array2   array2类型的数组
用法:
  str = "a=1&b=2"
  pat = "\w+=(\w+)"
  array2_match_php str,pat,arr
    '返回的为多维数组，结构大致如下：
    arr => (
      0 => (
        0 => a=1,
        1 => b=2,
      ),
      1 => (
        0 => 1,
        1 => 2,
      ),
    )


方法:array2_json_decode(arr,jss)
作用:将json字符串转换成arra2数组
参数:arr      array2数组
参数:jss      json字符串
返回:array2   array2类型的数组
用法:
  jss = "{a:1,b:2}"
  array2_json_decode arr,jss


方法:array2_json_encode(arr)
依赖:aspjson(<a href="https://code.google.com/p/aspjson/" target="_blank">点击获取</a>)
作用:将arra2数组转换成json字符串
参数:arr      array2数组
返回:string   json格式的字符串
用法:
  json = array2_json_encode(arr)


方法:array2_dump(arr)
作用:遍历打印array2数组，用于调试
参数:arr      array2数组
返回:string   遍历结果
用法:
  Response.Write array2_dump(arr)



</pre>