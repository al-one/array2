
此方法可以让Asp数组像PHP数组一样拥有键。

方法:array2(arr,byval k,byval v)
作用:创建/新增/修改数组元素
参数:arr      数组名
参数:k        键名key
参数:v        值value
用法:
  array2 arr,"key","hello would"  '创建一个名为arr的数组，并且有一个键为"key"、值为"hello would"的元素
  array2 arr,null,"some string"   '向数组arr中添加一个元素，值"some string"，当键为空/空字符串时，key自动为数组的所有key中最大一个数值+1，否则为0，此处key为0
  array2 arr,0,"hi Alone"         '修改数组中key为0的元素的值为"hi Alone"