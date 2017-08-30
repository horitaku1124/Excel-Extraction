import org.apache.poi.ss.usermodel.*
import java.io.BufferedWriter
import java.io.File
import java.nio.charset.Charset
import java.nio.file.Files
import java.nio.file.Paths
import java.time.ZoneId
import java.time.LocalDateTime
import java.time.format.DateTimeFormatter


val dateFormat: DateTimeFormatter = DateTimeFormatter.ofPattern("yyyy/MM/dd")
val dateTimeFormat: DateTimeFormatter = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss")
val timeFormat: DateTimeFormatter = DateTimeFormatter.ofPattern("HH:mm:ss")

fun main(args: Array<String>) {

  val csvEncode = Charset.forName("Shift_JIS")
  var sheetList = listOf("東京", "渋谷")

  val workbook = WorkbookFactory.create(File("./data/sample1.xlsx"))

  for(sheetName in sheetList) {
    val sheet = workbook.getSheet(sheetName)
    val csvPath = "./out/data/$sheetName.csv"
    println(csvPath)

    Files.newBufferedWriter(Paths.get(csvPath), csvEncode).use<BufferedWriter, Unit> {
      exportSheetToCsv(it, sheet)
    }
  }
}
fun exportSheetToCsv(writer:BufferedWriter, sheet: Sheet) {
  var headerCells =  mutableListOf<String>()
  var headerLimit = 0
  var row: Row = sheet.getRow(0) ?: return
  for (i in 0..50) {
    var cell: Cell? = row.getCell(i)
    if (cell == null) {
      headerCells.add("")
    } else {
      headerLimit = i
      headerCells.add(cellParseToString(cell))
    }
  }
  println("headerLimit=" + headerLimit)
  for (i in 0..10) {
    var row: Row? = sheet.getRow(i)
    var firstCell: Cell = row?.getCell(0) ?: break

    var lineCells =  mutableListOf<String>()
    for (j in 0 .. headerLimit) {
      var cell: Cell? = row.getCell(j)
      if (cell == null) {
        lineCells.add("")
        continue
      }

      println("$i, $j : " + cell.cellType)
      when (cell.cellType) {
        Cell.CELL_TYPE_NUMERIC -> {
          val numValue = cell.numericCellValue
          if (DateUtil.isCellDateFormatted(cell)) {
            val date = cell.dateCellValue
//            println("date? = " + numValue)

            val hasTime = (numValue - numValue.toInt().toDouble()) > 0.0
            val onlyTime = numValue < 1.0

            val localDateTime = LocalDateTime.ofInstant(date.toInstant(), ZoneId.systemDefault())
            if (onlyTime) {
              lineCells.add(timeFormat.format(localDateTime))
            } else if (hasTime) {
              lineCells.add(dateTimeFormat.format(localDateTime))
            } else {
              lineCells.add(dateFormat.format(localDateTime))
            }
          } else {
            lineCells.add(numValue.toString())
          }
        }
        Cell.CELL_TYPE_STRING -> lineCells.add(cell.stringCellValue)
        Cell.CELL_TYPE_FORMULA -> {
//          println(String.format("Formula:=%s cachedType:%d %s",
//                  cell.cellFormula,
//                  cell.cachedFormulaResultType,
//                  cell.numericCellValue.toString()
//          ))

          lineCells.add(cellParseToString(cell, cell.cachedFormulaResultType))
        }
      }
    }
    if (lineCells.isEmpty()) break
    writer.append(lineCells.joinToString(","))
    writer.append("\r\n")
  }
}

fun cellParseToString(cell: Cell): String {
  return cellParseToString(cell, null)
}
fun cellParseToString(cell: Cell, _type: Int?): String {
  var ret = ""
  val type = _type ?: cell.cellType
  when (type) {
    Cell.CELL_TYPE_NUMERIC -> ret = cell.numericCellValue.toString()
    Cell.CELL_TYPE_STRING -> ret = cell.stringCellValue
  }
  return ret
}