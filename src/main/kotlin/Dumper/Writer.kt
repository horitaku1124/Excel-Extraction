package Dumper

import org.apache.poi.ss.usermodel.Workbook

interface Writer: java.io.Closeable {
  fun header(title: String)
  fun write(s: List<List<String>>)
  fun sheetTitle(title: String)
  fun tableIndex(workbook: Workbook)
}