package cn.ljguo.android.ibrary.excel

import android.os.Build.VERSION_CODES.P
import android.util.Log
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.InputStream
import java.math.BigDecimal
import java.math.BigInteger
import java.time.LocalDate
import java.time.format.DateTimeFormatter
import java.util.*
import kotlin.collections.HashMap
import kotlin.reflect.KClass
import kotlin.reflect.KMutableProperty1
import kotlin.reflect.KProperty1
import kotlin.reflect.full.createInstance
import kotlin.reflect.full.declaredMemberProperties

class ExcelReader<T : Any> {
    private val tag = this.javaClass.name;
    private val propertyMap: HashMap<String, ExcelFiled<T>> = HashMap();
    fun read(inputStream: InputStream, kClazz: KClass<T>, pageReadListener: PageReadListener<T>) {
        val workbook = XSSFWorkbook(inputStream)
        val sheet = workbook.getSheetAt(0)
        pageReadListener.rowLength(sheet.lastRowNum)
        if (sheet.lastRowNum >= 2) {
            getClazzProperty(kClazz)
            getExcelProperty(sheet)
            for (r in 1 until sheet.lastRowNum) {
                val property = kClazz.createInstance()
                propertyMap.forEach {
                    val cell = sheet.getRow(r).getCell(it.value.number)
                    when (it.value.type.classifier) {
                        Int::class, java.lang.Integer::class -> it.value.property.set(
                            property,
                            cell.toString().toIntOrNull()
                        )
                        Boolean::class, java.lang.Boolean::class -> it.value.property.set(
                            property,
                            cell.toString().toBooleanStrictOrNull()
                        )
                        String::class, java.lang.String::class -> it.value.property.set(
                            property,
                            cell.toString()
                        )
                        Long::class, java.lang.Long::class -> it.value.property.set(
                            property,
                            cell.toString().toLongOrNull()
                        )
                        Double::class, java.lang.Double::class -> it.value.property.set(
                            property,
                            cell.toString().toDoubleOrNull()
                        )
                        Float::class, java.lang.Float::class -> it.value.property.set(
                            property,
                            cell.toString().toFloatOrNull()
                        )
                        BigDecimal::class -> it.value.property.set(
                            property,
                            cell.toString().toBigDecimalOrNull()
                        )
                        BigInteger::class -> it.value.property.set(
                            property,
                            cell.toString().toBigIntegerOrNull()
                        )
                        Byte::class, java.lang.Byte::class -> it.value.property.set(
                            property,
                            cell.toString().toByteOrNull()
                        )
                        Date::class -> {
                            if (cell.cellType.equals(CellType.NUMERIC)) {
                                it.value.property.set(property, cell.dateCellValue)
                            } else {
                                it.value.property.set(
                                    property,
                                    LocalDate.parse(
                                        cell.toString(),
                                        DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss")
                                    )
                                )
                            }
                        }
                        else -> it.value.property.set(property, cell.toString())
                    }
                }
                pageReadListener.read(property, r);
                Log.d(tag, property.toString());
            }
        }
        pageReadListener.end()
    }

    private fun getExcelProperty(sheet: Sheet) {
        val headRow = sheet.getRow(0);
        for (c in 0 until headRow.lastCellNum) {
            val cell = headRow.getCell(c);
            if (propertyMap[cell.toString()] != null) {
                if (propertyMap[cell.toString()]!!.number == 0) {
                    propertyMap[cell.toString()]!!.number = c
                }
            }

        }
    }

    private fun getClazzProperty(kClazz: KClass<T>) {
        kClazz.declaredMemberProperties.forEach { property ->
            val annotationList = property.annotations;
            var isExcelPropertyAnnotation = false
            if (annotationList.isNotEmpty()) {
                annotationList.forEach { annotation ->
                    if (annotation is ExcelProperty) {
                        propertyMap[annotation.value] = ExcelFiled(
                            0,
                            property.name,
                            property.returnType,
                            property as KProperty1<T, Any?> as KMutableProperty1<T, Any?>
                        )
                        isExcelPropertyAnnotation = true
                    }
                }
            }
            if (!isExcelPropertyAnnotation) {
                propertyMap[property.name] = ExcelFiled(
                    0,
                    property.name,
                    property.returnType,
                    property as KProperty1<T, Any?> as KMutableProperty1<T, Any?>
                )
            }
        }
    }
}