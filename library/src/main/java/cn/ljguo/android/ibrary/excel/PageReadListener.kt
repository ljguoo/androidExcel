package cn.ljguo.android.ibrary.excel

interface PageReadListener<T> {
   fun rowLength(length:Int)
   fun read(t:T,number:Int)
   fun end()
}