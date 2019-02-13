package util

import org.apache.poi.ss.format.CellFormat
import org.apache.poi.ss.usermodel.*
import java.time.LocalDateTime
import java.time.ZoneId
import java.time.format.DateTimeFormatter
import java.util.regex.Pattern

class ExcelUtil {
  companion object {
    private val dateFormat: DateTimeFormatter = DateTimeFormatter.ofPattern("yyyy/MM/dd")
    private val dateTimeFormat: DateTimeFormatter = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss")
    private val timeFormat: DateTimeFormatter = DateTimeFormatter.ofPattern("HH:mm:ss")

    private val intNumPattern: Pattern = Pattern.compile("([\\d-]+)\\.(\\d+)E(\\d+)")

    fun cellParseToString(cell: Cell, type: Int = cell.cellType): String {
      return when (type) {
        Cell.CELL_TYPE_NUMERIC -> {
          val numValue = cell.numericCellValue
          var numString = numValue.toString()

          if(BuiltinFormats.FIRST_USER_DEFINED_FORMAT_INDEX <= cell.cellStyle.dataFormat) {
            val cellFormat = CellFormat.getInstance(cell.cellStyle.dataFormatString)
            numString = cellFormat.apply(cell).text
          } else if (DateUtil.isCellDateFormatted(cell)) {
            val date = cell.dateCellValue

            val hasTime = (numValue - numValue.toInt().toDouble()) > 0.0
            val onlyTime = numValue < 1.0

            val localDateTime = LocalDateTime.ofInstant(date.toInstant(), ZoneId.systemDefault())
            numString = when {
              cell.cellStyle.dataFormat.toInt() == 22 -> {
                val cellFormat = CellFormat.getInstance("yyyy/mm/dd\\ h:mm")
                cellFormat.apply(cell).text
              }
              onlyTime -> timeFormat.format(localDateTime)
              hasTime -> dateTimeFormat.format(localDateTime)
              else -> dateFormat.format(localDateTime)
            }
          } else {
            // Number
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
                val floatPoints = number2.length - digit
                numString = String.format("%$digit." + floatPoints + "f", numValue)
              }
            } else if(numString.endsWith(".0")) {
              numString = numString.substring(0, numString.length - 2)
            }
          }
          numString
        }
        Cell.CELL_TYPE_STRING -> cell.stringCellValue
        else -> ""
      }
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
  }
}