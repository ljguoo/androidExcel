package cn.ljguo.android.excel.bean

import cn.ljguo.android.ibrary.excel.ExcelProperty
import java.util.*

data class Test(
    @ExcelProperty(value = "姓名")
    var name:String?="",
    @ExcelProperty(value = "年龄")
    var age:String?="",
    @ExcelProperty(value = "日期")
    var date:Date?=null
)
