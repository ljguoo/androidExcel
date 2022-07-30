### 安卓读取或写入excel

目前只实现了读取Excel

1. 新建实体类

   ```kotlin
   data class Test(
       @ExcelProperty(value = "姓名")
       var name:String?="",
       @ExcelProperty(value = "年龄")
       var age:String?="",
       @ExcelProperty(value = "日期")
       var date:Date?=null
   )
   ```

 2.读取文件

```kotlin
ExcelReader<Test>().read(this.resources.openRawResource(R.raw.test),Test::class, object :
    PageReadListener<Test> {
    override fun rowLength(length: Int) {
        Log.d(tag, "excel长度=${length}")
    }

    override fun read(t: Test, number: Int) {
        Log.d(tag, "第${number}行的数据=${t}")
    }

    override fun end() {
        Log.d(tag, "解析结束")
    }

})
```