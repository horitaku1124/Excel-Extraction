import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.WorkbookFactory
import util.ExcelUtil
import java.io.FileInputStream
import java.util.regex.Pattern

class ExcelQuery {
  companion object {
    private val intNumPattern: Pattern = Pattern.compile("select\\s+(\\S+)\\s+from\\s+(`?)([/.a-zA-Z0-9]+)(`?)\\.?(\\S*)")
    @JvmStatic
    fun main(args: Array<String>) {
      if (args.isEmpty()) {
        error("no argument")
      }
      val query = args[0]
      val matcher = intNumPattern.matcher(query)
      if (!matcher.find()) {
        error("syntax error")
      }
      val selector = matcher.group(1)
      var bq1 = matcher.group(2)
      val schema = matcher.group(3)
      var bq2 = matcher.group(4)
      val sheetName = matcher.group(5)
//      println("selector=${selector}")
//      println("schema=${schema}")
//      println("sheetName=${sheetName}")

      val workbook = WorkbookFactory.create(FileInputStream(schema))
      val sheet = workbook.getSheet(sheetName)
      val header = ExcelUtil.fetchHeader(sheet)

      var selectRange = mutableListOf<Int>()

      if (selector == "*" ) {
        for (i in 0 until header.size) {
          selectRange.add(i)
        }
      } else {
        selector.split(",").forEach {
          var found = false
          for (i in 0 until header.size) {
            if (header[i] == it) {
              selectRange.add(i)
              found = true
              break
            }
          }
          if (!found) {
            error("No elemet => ${it}")
          }
        }
      }

      println(selectRange.map({header[it]}).joinToString(","))

      for (i in 1 .. 1000) {
        val dataRow: Row? = sheet.getRow(i)
        dataRow?.getCell(0) ?: break

        val lineCells = mutableListOf<String>()
        for (j in selectRange) {
          val cell: Cell? = dataRow.getCell(j)
          if (cell == null) {
            lineCells.add("")
            continue
          }

          lineCells.add(
              when (cell.cellType) {
                Cell.CELL_TYPE_NUMERIC -> ExcelUtil.cellParseToString(cell)
                Cell.CELL_TYPE_STRING -> cell.stringCellValue
                Cell.CELL_TYPE_FORMULA -> ExcelUtil.cellParseToString(cell, cell.cachedFormulaResultType)
                else -> ""
              }
          )
        }
        if (lineCells.isEmpty()) break
        println(lineCells.joinToString(","))
      }
    }
  }
}
