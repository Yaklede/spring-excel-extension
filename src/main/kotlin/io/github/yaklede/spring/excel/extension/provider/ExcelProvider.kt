package io.github.yaklede.spring.excel.extension.provider

import io.github.yaklede.spring.excel.extension.annotation.ExcelColumn
import com.empty.generator.EmptyObjectGenerator
import jakarta.servlet.http.HttpServletResponse

import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.net.URLEncoder
import java.time.LocalDate
import java.time.LocalDateTime
import kotlin.reflect.full.declaredMemberProperties
import kotlin.reflect.full.findAnnotation
import kotlin.reflect.full.memberProperties

/**
 * [List]를 엑셀 파일로 변경합니다.
 *
 * 엑셀에 표현될 프로퍼티는 모두 [ExcelColumn]을 사용하여 컬럼 명과 인덱스를 지정해주어야 합니다.
 *
 * @param fileName 생성될 엑셀 파일 명 ex) 종합_리포트
 * @param response IO 대상
 */
inline fun <reified T : Any> List<T>.toExcelFile(
    fileName: String,
    response: HttpServletResponse,
): HttpServletResponse {

    val workbook = XSSFWorkbook() // 엑셀 파일 객체 생성
    val sheet: XSSFSheet = workbook.createSheet() // 시트 생성

    val headerList: Map<Int, String> = this.let { list ->
        if (list.isEmpty())
            EmptyObjectGenerator.generate(T::class, isNullable = true)
        else
            this.first()
    }.getAnnotatedMemberPropertyNames()

    // 셀 가로 길이 설정
    headerList.forEach { entry ->
        sheet.apply { setColumnWidth(entry.key + 1, 20 * 256) }
    }

    // 스타일 설정
    val defaultCellStyle = defaultStyle(workbook)

    // 컬럼 헤더 설정
    val headerRow = sheet.createRow(0)
    with(headerRow) {
        createCell(0).apply {
            setCellValue("No")
            cellStyle = defaultCellStyle
        }
        headerList.forEach { (key, _) ->
            createCell(key + 1).apply {
                setCellValue(headerList[key])
                cellStyle = defaultCellStyle
            }
        }
    }

    val dateTimeStyle = dateTimeStyle(workbook)
    val dateStyle = dateStyle(workbook)
    val numberStyle = numberStyle(workbook)

    this.forEachIndexed { index, any ->
        val row = sheet.createRow(index + 1)

        row.createCell(0).apply {
            cellStyle = defaultCellStyle
            setCellValue(index + 1.0)
        }

        val properties = any.getAnnotatedMemberPropertyValues()

        properties.asSequence().forEach { (key, value) ->
            val cell = row.createCell(key + 1).apply {
                cellStyle = defaultCellStyle
            }
            when (value) {
                is String -> cell.apply {
                    setCellValue(value)
                }

                is Number -> cell.apply {
                    cellStyle = numberStyle
                    setCellValue(value.toDouble())
                }

                is LocalDateTime -> cell.apply {
                    cellStyle = dateTimeStyle
                    setCellValue(value)
                }

                is LocalDate -> cell.apply {
                    cellStyle = dateStyle
                    setCellValue(value)
                }

                is Enum<*> -> cell.apply {
                    val enum = value::class.memberProperties.map { kProperty1 ->
                        kProperty1.call(value)
                    }[0] as String
                    setCellValue(enum)
                }

                else -> cell.apply {
                    setCellValue(value.toString())
                }
            }
        }
    }

    response.contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    response.addHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(fileName, "UTF-8") + ".xlsx")
    workbook.write(response.outputStream)
    workbook.close()

    return response
}

inline fun <reified T : Any> T.getAnnotatedMemberPropertyNames(): MutableMap<Int, String> {

    val excelColumns = this::class.declaredMemberProperties
        .mapNotNull { kProperty1 ->
            kProperty1.findAnnotation<ExcelColumn>()
        }

    val headerNames = mutableMapOf<Int, String>()
    excelColumns.forEach {
        headerNames[it.index] = it.name
    }

    return headerNames
}

inline fun <reified T : Any> T.getAnnotatedMemberPropertyValues(): Map<Int, Any> {

    val memberProperties = this::class.declaredMemberProperties

    val properties = mutableMapOf<Int, Any>()
    memberProperties.forEach { kProperty1 ->
        val key = kProperty1.findAnnotation<ExcelColumn>()?.index
        val value = kProperty1.call(this) ?: ""

        if (key != null)
            properties[key] = value
    }

    return properties
}

fun defaultStyle(workbook: XSSFWorkbook): XSSFCellStyle =
    workbook.createCellStyle().apply {
        borderTop = BorderStyle.THIN
        borderBottom = BorderStyle.THIN
        borderLeft = BorderStyle.THIN
        borderRight = BorderStyle.THIN
    }

/**
 * 일시 포맷 스타일 적용
 *
 * @param workbook
 */
fun dateTimeStyle(workbook: XSSFWorkbook): XSSFCellStyle =
    workbook.createCellStyle().apply {
        dataFormat = workbook.creationHelper.createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss")
        borderTop = BorderStyle.THIN
        borderBottom = BorderStyle.THIN
        borderLeft = BorderStyle.THIN
        borderRight = BorderStyle.THIN
    }

/**
 * 일자 포맷 스타일 적용
 *
 * @param workbook
 */
fun dateStyle(workbook: XSSFWorkbook): XSSFCellStyle =
    workbook.createCellStyle().apply {
        dataFormat = workbook.creationHelper.createDataFormat().getFormat("yyyy-MM-dd")
        borderTop = BorderStyle.THIN
        borderBottom = BorderStyle.THIN
        borderLeft = BorderStyle.THIN
        borderRight = BorderStyle.THIN
    }

/**
 * 숫자 포맷 스타일 적용
 *
 * @param workbook
 */
fun numberStyle(workbook: XSSFWorkbook): XSSFCellStyle =
    workbook.createCellStyle().apply {
        dataFormat = workbook.creationHelper.createDataFormat().getFormat("#,##0")
        borderTop = BorderStyle.THIN
        borderBottom = BorderStyle.THIN
        borderLeft = BorderStyle.THIN
        borderRight = BorderStyle.THIN
    }
