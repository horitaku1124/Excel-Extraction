import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.WorkbookFactory
import util.ExcelUtil
import java.io.BufferedWriter
import java.io.File
import java.io.FileInputStream
import java.nio.charset.Charset
import java.nio.file.Files
import java.nio.file.Paths

class ExtractBySheet {
  companion object {
    private val csvEncode: Charset = Charset.forName("Shift_JIS")


    @JvmStatic
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

    private fun exportSheetToCsv(writer:BufferedWriter, sheet: Sheet, config: Configuration) {
      val headerCells = ExcelUtil.fetchHeader(sheet)

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
                Cell.CELL_TYPE_NUMERIC -> ExcelUtil.cellParseToString(cell)
                Cell.CELL_TYPE_STRING -> cell.stringCellValue
                Cell.CELL_TYPE_FORMULA -> ExcelUtil.cellParseToString(cell, cell.cachedFormulaResultType)
                else -> ""
              }
          )
        }
        if (lineCells.isEmpty()) break
        writer.append(lineCells.joinToString(","))
        writer.append("\r\n")
        len++
      }
      println("rows=$len")
    }


    private fun exportSheetToCsvDivided(csvPathFormat:String, sheet: Sheet, divideItem: Int, config: Configuration) {
      var writeFile: File? = null;
      var writer: BufferedWriter? = null

      val headerCells = ExcelUtil.fetchHeader(sheet)
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
                Cell.CELL_TYPE_NUMERIC -> ExcelUtil.cellParseToString(cell)
                Cell.CELL_TYPE_STRING -> cell.stringCellValue
                Cell.CELL_TYPE_FORMULA -> ExcelUtil.cellParseToString(cell, cell.cachedFormulaResultType)
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
      println("rows=$len")
    }


  }
}