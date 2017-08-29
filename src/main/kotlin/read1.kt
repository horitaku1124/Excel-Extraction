import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.BufferedWriter
import java.io.File
import java.nio.charset.Charset
import java.nio.file.Files
import java.nio.file.Paths


fun main(args: Array<String>) {
  var sheetList = listOf("東京")

  val workbook = WorkbookFactory.create(File("./data/sample1.xlsx"))

  for(sheetName in sheetList) {
    val sheet = workbook.getSheet(sheetName)
    val csvPath:String = "./out/data/" + sheetName + ".csv"
    println(csvPath)

    Files.newBufferedWriter(Paths.get(csvPath), Charset.forName("Shift_JIS")).use<BufferedWriter, Unit> {
      for (i in 0..10) {
        var row: Row? = sheet.getRow(i)
        var firstCell: Cell = row?.getCell(0) ?: break

        var lineCells =  mutableListOf<String>()
        for (j in 0..10) {
          var cell: Cell = row.getCell(j) ?: break

          if(cell.cellType == Cell.CELL_TYPE_NUMERIC){
            lineCells.add(cell.numericCellValue.toString())
          }
          if(cell.cellType == Cell.CELL_TYPE_STRING){
            lineCells.add(cell.stringCellValue)
          }
        }
        if (lineCells.isEmpty()) break
        it.append(lineCells.joinToString(","))
        it.append("\r\b")
      }
    }
  }
}