import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.DateUtil
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.BufferedWriter
import java.io.File
import java.nio.charset.Charset
import java.nio.file.Files
import java.nio.file.Paths
import java.time.ZoneId
import java.time.LocalDateTime
import java.time.format.DateTimeFormatter




fun main(args: Array<String>) {
  val dateFormat = DateTimeFormatter.ofPattern("yyyy/MM/dd")
  val dateTimeFormat = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss.SSS")

  val csvEncode = Charset.forName("Shift_JIS")
  var sheetList = listOf("東京")

  val workbook = WorkbookFactory.create(File("./data/sample1.xlsx"))

  for(sheetName in sheetList) {
    val sheet = workbook.getSheet(sheetName)
    val csvPath = "./out/data/$sheetName.csv"
    println(csvPath)

    Files.newBufferedWriter(Paths.get(csvPath), csvEncode).use<BufferedWriter, Unit> {
      for (i in 0..10) {
        var row: Row? = sheet.getRow(i)
        var firstCell: Cell = row?.getCell(0) ?: break

        var lineCells =  mutableListOf<String>()
        for (j in 0..10) {
          var cell: Cell = row.getCell(j) ?: break

          println("$i, $j : " + cell.cellType)
          when (cell.cellType) {
            Cell.CELL_TYPE_NUMERIC -> {
              if (DateUtil.isCellDateFormatted(cell)) {
                val date = cell.dateCellValue

                val localDateTime = LocalDateTime.ofInstant(date.toInstant(), ZoneId.systemDefault())
                lineCells.add(dateFormat.format(localDateTime))
              } else {
                lineCells.add(cell.numericCellValue.toString())
              }
            }
            Cell.CELL_TYPE_STRING -> lineCells.add(cell.stringCellValue)
            Cell.CELL_TYPE_FORMULA -> {
              println(String.format("Formula:=%s cachedType:%d %s",
                  cell.cellFormula,
                  cell.cachedFormulaResultType,
                  cell.numericCellValue.toString()
              ))

              when (cell.cachedFormulaResultType) {
                Cell.CELL_TYPE_NUMERIC -> lineCells.add(cell.numericCellValue.toString())
                Cell.CELL_TYPE_STRING -> lineCells.add(cell.stringCellValue)
              }
            }
          }
        }
        if (lineCells.isEmpty()) break
        it.append(lineCells.joinToString(","))
        it.append("\r\n")
      }
    }
  }
}