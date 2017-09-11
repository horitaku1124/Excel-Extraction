import org.apache.poi.ss.format.CellFormat
import org.apache.poi.ss.usermodel.*
import java.io.BufferedWriter
import java.io.File
import java.io.FileInputStream
import java.nio.charset.Charset
import java.nio.file.Files
import java.nio.file.Paths
import java.time.LocalDateTime
import java.time.ZoneId
import java.time.format.DateTimeFormatter
import java.util.regex.Pattern


val dateFormat: DateTimeFormatter = DateTimeFormatter.ofPattern("yyyy/MM/dd")
val dateTimeFormat: DateTimeFormatter = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss")
val timeFormat: DateTimeFormatter = DateTimeFormatter.ofPattern("HH:mm:ss")
val csvEncode: Charset = Charset.forName("Shift_JIS")

val intNumPattern: Pattern = Pattern.compile("([\\d-]+)\\.(\\d+)E(\\d+)")
fun main(args: Array<String>) {
  val config = Configuration(args)
  val sheetList = config.sheets
  val divideItems = config.divideItems
  val outputDirectory = config.outputDirectory

  if (outputDirectory.isBlank()) {
    println("outputDirectory is blank")
    System.exit(1)
  }

  val workbook = WorkbookFactory.create(FileInputStream(config.inputFile))

  for(sheetName in sheetList) {
    val sheet = workbook.getSheet(sheetName)

    if (divideItems > 1) {
      val csvPathFormat = outputDirectory + "/$sheetName" + "_%d.csv"
      exportSheetToCsvDivided(csvPathFormat, sheet, divideItems, config)
    } else {
      val csvPath = "$outputDirectory/$sheetName.csv"
      println(csvPath)
      Files.newBufferedWriter(Paths.get(csvPath), csvEncode).use<BufferedWriter, Unit> {
        exportSheetToCsv(it, sheet, config)
      }
    }
  }
}
fun exportSheetToCsv(writer:BufferedWriter, sheet: Sheet, config: Configuration) {
  val headerCells = fetchHeader(sheet)

  writer.append(headerCells.joinToString(","))
  writer.append("\r\n")

  println("headerLimit=" + headerCells.size)
  var len = 0
  for (i in 1 .. config.limit) {
    val dataRow: Row? = sheet.getRow(i)
    dataRow?.getCell(0) ?: break

    val lineCells =  mutableListOf<String>()
    for (j in 0 .. headerCells.size) {
      val cell: Cell? = dataRow.getCell(j)
      if (cell == null) {
        lineCells.add("")
        continue
      }

      lineCells.add(
          when (cell.cellType) {
            Cell.CELL_TYPE_NUMERIC -> cellParseToString(cell)
            Cell.CELL_TYPE_STRING -> cell.stringCellValue
            Cell.CELL_TYPE_FORMULA -> cellParseToString(cell, cell.cachedFormulaResultType)
            else -> ""
          }
      )
    }
    if (lineCells.isEmpty()) break
    writer.append(lineCells.joinToString(","))
    writer.append("\r\n")
    len++
  }
  println("rows=" + len)
}


fun exportSheetToCsvDivided(csvPathFormat:String, sheet: Sheet, divideItem: Int, config: Configuration) {
  var writeFile: File? = null;
  var writer: BufferedWriter? = null

  val headerCells = fetchHeader(sheet)
  println("headerLimit=" + headerCells.size)

  var len = 0
  var fileOffset = 0
  var lenOnFile = 0
  for (i in 1 .. config.limit) {
    if (writer == null) {
      val csvPath = String.format(csvPathFormat, fileOffset)
      println(csvPath)
      writeFile = File(csvPath)
      writer = writeFile.bufferedWriter(csvEncode)

      writer.append(headerCells.joinToString(","))
      writer.append("\r\n")
      lenOnFile = 0
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

      lineCells.add(
          when (cell.cellType) {
            Cell.CELL_TYPE_NUMERIC -> cellParseToString(cell)
            Cell.CELL_TYPE_STRING -> cell.stringCellValue
            Cell.CELL_TYPE_FORMULA -> cellParseToString(cell, cell.cachedFormulaResultType)
            else -> ""
          }
      )
    }
    if (lineCells.isEmpty()) break
    writer.append(lineCells.joinToString(","))
    writer.append("\r\n")
    len++
    lenOnFile++
    if(len % divideItem == 0) {
      writer.close()
      writer = null
      fileOffset++
    }
  }
  if (writer != null) {
    writer.close()
  }
  if (writeFile != null && lenOnFile == 0) {
    writeFile.delete()
  }
  println("rows=" + len)
}

fun fetchHeader(sheet: Sheet) : MutableList<String>{
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

fun cellParseToString(cell: Cell, type: Int = cell.cellType): String {
  return when (type) {
    Cell.CELL_TYPE_NUMERIC -> {
      val numValue = cell.numericCellValue
      var numString = numValue.toString()
      if (DateUtil.isCellDateFormatted(cell)) {
        val date = cell.dateCellValue

        val hasTime = (numValue - numValue.toInt().toDouble()) > 0.0
        val onlyTime = numValue < 1.0

        val localDateTime = LocalDateTime.ofInstant(date.toInstant(), ZoneId.systemDefault())
        numString = when {
          BuiltinFormats.FIRST_USER_DEFINED_FORMAT_INDEX <= cell.cellStyle.dataFormat -> {
            val cellFormat = CellFormat.getInstance(cell.cellStyle.dataFormatString)
            cellFormat.apply(cell).text
          }
          cell.cellStyle.dataFormat.toInt() == 22 -> {
            val cellFormat = CellFormat.getInstance("yyyy/mm/dd\\ h:mm")
            cellFormat.apply(cell).text
          }
          onlyTime -> timeFormat.format(localDateTime)
          hasTime -> dateTimeFormat.format(localDateTime)
          else -> dateFormat.format(localDateTime)
        }
//        println("numString=" + numString)
//        println(" dataFormat=" + cell.cellStyle.dataFormat)
//        println(" dataFormatString=" + cell.cellStyle.dataFormatString)
      } else {
        // Number
//        println("numString=" + numString)
        val matcher = intNumPattern.matcher(numString)
        if (matcher.find()) {
          val number1 = matcher.group(1)
          val number2 = matcher.group(2)
          val digit = matcher.group(3).toInt()
          if (number2.length == digit) {
            numString = number1 + number2
          } else if (number2.length < digit) {
            numString = String.format("%$digit.0f", numValue)
          } else {
            var floatPoints = number2.length - digit
            numString = String.format("%$digit." + floatPoints + "f", numValue)
          }
        }
//        println(" numString2=" + numString)
      }
      numString
    }
    Cell.CELL_TYPE_STRING -> cell.stringCellValue
    else -> ""
  }
}