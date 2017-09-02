import org.apache.poi.ss.format.CellFormat
import org.apache.poi.ss.usermodel.*
import java.io.BufferedWriter
import java.io.FileInputStream
import java.nio.charset.Charset
import java.nio.file.Files
import java.nio.file.Paths
import java.time.LocalDateTime
import java.time.ZoneId
import java.time.format.DateTimeFormatter


val dateFormat: DateTimeFormatter = DateTimeFormatter.ofPattern("yyyy/MM/dd")
val dateTimeFormat: DateTimeFormatter = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss")
val timeFormat: DateTimeFormatter = DateTimeFormatter.ofPattern("HH:mm:ss")
val csvEncode: Charset = Charset.forName("Shift_JIS")

fun main(args: Array<String>) {
  val sheetList = listOf("新宿")
  val divideItems = 3

  val workbook = WorkbookFactory.create(FileInputStream("./data/sample1.xlsx"))

  for(sheetName in sheetList) {
    val sheet = workbook.getSheet(sheetName)

    if (divideItems > 1) {
      val csvPathFormat = "./out/data/$sheetName" + "_%d.csv"
      exportSheetToCsvDivided(csvPathFormat, sheet, divideItems)
    } else {
      val csvPath = "./out/data/$sheetName.csv"
      println(csvPath)
      Files.newBufferedWriter(Paths.get(csvPath), csvEncode).use<BufferedWriter, Unit> {
        exportSheetToCsv(it, sheet)
      }
    }
  }
}
fun exportSheetToCsv(writer:BufferedWriter, sheet: Sheet) {
  val headerCells = fetcHeader(sheet)

  writer.append(headerCells.joinToString(","))
  writer.append("\r\n")

  println("headerLimit=" + headerCells.size)
  var len = 0
  for (i in 1 .. Int.MAX_VALUE) {
    val dataRow: Row? = sheet.getRow(i)
    dataRow?.getCell(0) ?: break

    val lineCells =  mutableListOf<String>()
    for (j in 0 .. headerCells.size) {
      val cell: Cell? = dataRow.getCell(j)
      if (cell == null) {
        lineCells.add("")
        continue
      }

      when (cell.cellType) {
        Cell.CELL_TYPE_NUMERIC -> {
          lineCells.add(cellParseToString(cell))
        }
        Cell.CELL_TYPE_STRING -> lineCells.add(cell.stringCellValue)
        Cell.CELL_TYPE_FORMULA -> {
          lineCells.add(cellParseToString(cell, cell.cachedFormulaResultType))
        }
      }
    }
    if (lineCells.isEmpty()) break
    writer.append(lineCells.joinToString(","))
    writer.append("\r\n")
    len++
  }
  println("rows=" + len)
}


fun exportSheetToCsvDivided(csvPathFormat:String, sheet: Sheet, divideItem: Int) {
  var writer: BufferedWriter? = null

  val headerCells = fetcHeader(sheet)
  println("headerLimit=" + headerCells.size)

  var len = 0
  var fileOffset = 0
  for (i in 1 .. Int.MAX_VALUE) {
    if (writer == null) {
      var csvPath = String.format(csvPathFormat, fileOffset)
      println(csvPath)
      writer = Files.newBufferedWriter(Paths.get(csvPath), csvEncode)

      writer.append(headerCells.joinToString(","))
      writer.append("\r\n")
    }
    val dataRow: Row? = sheet.getRow(i)
    dataRow?.getCell(0) ?: break

    val lineCells =  mutableListOf<String>()
    for (j in 0 .. headerCells.size) {
      val cell: Cell? = dataRow.getCell(j)
      if (cell == null) {
        lineCells.add("")
        continue
      }

      when (cell.cellType) {
        Cell.CELL_TYPE_NUMERIC -> {
          lineCells.add(cellParseToString(cell))
        }
        Cell.CELL_TYPE_STRING -> lineCells.add(cell.stringCellValue)
        Cell.CELL_TYPE_FORMULA -> {
          lineCells.add(cellParseToString(cell, cell.cachedFormulaResultType))
        }
      }
    }
    if (lineCells.isEmpty()) break
    writer?.append(lineCells.joinToString(","))
    writer?.append("\r\n")
    len++
    if(len % divideItem == 0) {
      writer?.close()
      writer = null
      fileOffset++
    }
  }
  if (writer != null) {
    writer.close()
  }
  println("rows=" + len)
}

fun fetcHeader(sheet: Sheet) : MutableList<String>{
  val headerCells = mutableListOf<String>()
  var headerLimit = 0
  val row: Row = sheet.getRow(0) ?: return headerCells
  for (i in 0..50) {
    val cell: Cell? = row.getCell(i)
    if (cell == null) {
      headerCells.add("")
    } else {
      headerLimit = i
      headerCells.add(cellParseToString(cell))
    }
  }
  val retHeaderCells = mutableListOf<String>()
  for(i in 0..headerLimit) {
    retHeaderCells.add(headerCells.get(i))
  }
  return retHeaderCells
}

fun cellParseToString(cell: Cell): String {
  return cellParseToString(cell, null)
}
fun cellParseToString(cell: Cell, _type: Int?): String {
  var ret = ""
  val type = _type ?: cell.cellType
  when (type) {
    Cell.CELL_TYPE_NUMERIC -> {
      val numValue = cell.numericCellValue
      var numString = numValue.toString()
      if (DateUtil.isCellDateFormatted(cell)) {
        val date = cell.dateCellValue

        val hasTime = (numValue - numValue.toInt().toDouble()) > 0.0
        val onlyTime = numValue < 1.0

        val localDateTime = LocalDateTime.ofInstant(date.toInstant(), ZoneId.systemDefault())
        if (onlyTime) {
          numString = timeFormat.format(localDateTime)
        } else if (hasTime) {
          numString = dateTimeFormat.format(localDateTime)
        } else {
          numString = dateFormat.format(localDateTime)
        }
        if (BuiltinFormats.FIRST_USER_DEFINED_FORMAT_INDEX <= cell.cellStyle.dataFormat) {
          val cellFormat = CellFormat.getInstance(cell.cellStyle.dataFormatString)
          val cellFormatResult = cellFormat.apply(cell)
          numString = cellFormatResult.text
        }
      }
      return numString
    }
    Cell.CELL_TYPE_STRING -> ret = cell.stringCellValue
  }
  return ret
}