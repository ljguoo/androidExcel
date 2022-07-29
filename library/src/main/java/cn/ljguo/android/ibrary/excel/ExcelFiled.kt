package cn.ljguo.android.ibrary.excel

import kotlin.reflect.KMutableProperty1
import kotlin.reflect.KType

data class ExcelFiled<T>(var number:Int, val name:String, val type: KType, val property: KMutableProperty1<T, Any?>)