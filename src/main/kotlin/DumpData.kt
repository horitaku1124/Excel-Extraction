import Dumper.HtmlWriter
import Dumper.Writer
import org.apache.poi.ss.usermodel.WorkbookFactory
import util.ExcelUtil.Companion.cellToStringSimple
import java.io.FileInputStream

class DumpData {
  companion object {
    @JvmStatic
    fun main(args: Array<String>) {
      if (args.isEmpty()) {
        error("no argument")
      }
      var dumpFile = args[0]
      val workbook = WorkbookFactory.create(FileInputStream(dumpFile))

      var writer: Writer = HtmlWriter(System.out)
      writer.use {
        for (sheet in workbook.sheetIterator()) {
          it.sheetTitle(sheet.sheetName)
          var list = mutableListOf<MutableList<String>>()
          for (row in sheet.rowIterator()) {
            val cells = mutableListOf<String>()
            for (cell in row.cellIterator()) {
              cells.add(cellToStringSimple(cell))
            }
            list.add(cells)
          }
          it.write(list)
        }
      }
    }
  }
}