package Dumper

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
    fout.write("\n\n<h1>$title</h1>\n".toByteArray())
  }

  override fun close() {
    fout.write("</body>\n</html>\n".toByteArray())
    fout.close()
  }

  fun po(c: Int, fo: OutputStream) {
    if (c > 1) {
      fo.write("  <td colspan='$c'></td>\n".toByteArray())
    } else {
      fo.write("  <td></td>\n".toByteArray())
    }
  }

  override fun write(data: List<List<String>>) {
    var limit = 0
    data.forEach {
      if (limit < it.size) {
        limit = it.size
      }
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