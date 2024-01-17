package io.github.yaklede.spring.excel.extension.annotation

@Target(AnnotationTarget.PROPERTY)
@Retention(AnnotationRetention.RUNTIME)
annotation class ExcelColumn(
    val name: String,
    val index: Int = 0,
)

