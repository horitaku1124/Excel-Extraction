package Dumper

import org.apache.poi.ss.usermodel.Workbook
import java.io.OutputStream

class HtmlWriter(var fout: OutputStream) : Writer{

  init {
    fout.write("""<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
</head>
<body>
""".toByteArray())
  }
  override fun sheetTitle(title: String) {
    fout.write("\n\n<h1><a name='$title'>$title</a></h1>\n".toByteArray())
  }

  override fun close() {
    fout.write("</body>\n</html>\n".toByteArray())
    fout.close()
  }

  private fun po(c: Int, fo: OutputStream) {
    if (c > 1) {
      fo.write("  <td colspan='$c'></td>\n".toByteArray())
    } else {
      fo.write("  <td></td>\n".toByteArray())
    }
  }
  private fun escape(str:String):String {
    return str
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace("\n", "<br>\n")
  }
  override fun tableIndex(workbook: Workbook) {
    var buf = StringBuffer()
    buf.append("<div>\n")
    for (sheet in workbook.sheetIterator()) {
      var sheetName = escape(sheet.sheetName)
      buf.append("<p>")
      buf.append("<a href='#$sheetName'>")
      buf.append(sheetName)
      buf.append("</a>")
      buf.append("</p>\n")
    }
    buf.append("</div>\n")
    fout.write(buf.toString().toByteArray())
  }

  override fun write(data: List<List<String>>) {
    var limit = 0
    data.forEach {
      var right = it.size - 1
      while(limit < right) {
        if (it[right].isNotEmpty()) {
          limit = right
          break
        }
        right--
      }
    }

    fout.write("<table>\n".toByteArray())
    data.forEach {
      fout.write(" <tr>\n".toByteArray())

      var compound = 0
      for (i in 0..limit) {
        var cell = ""
        if (i < it.size) {
          cell = it[i]
        }

        if (cell == "") {
          compound++
        } else {
          if (compound > 0) {
            po(compound, fout)
            compound = 0
          }
          cell = escape(cell)
          fout.write("  <td>$cell</td>\n".toByteArray())
        }
      }
      if (compound > 0) {
        po(compound, fout)
      }
      fout.write(" </tr>\n".toByteArray())
    }
    fout.write("</table>\n".toByteArray())
  }

  override fun header(title: String) {
  }

}