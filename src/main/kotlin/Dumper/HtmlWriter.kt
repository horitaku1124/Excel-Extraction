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
    fout.write("<h1>$title</h1>\n".toByteArray())
  }

  override fun close() {
    fout.write("</body>\n</html>\n".toByteArray())
    fout.close()
  }

  override fun write(data: List<List<String>>) {
    var limit = 0
    data.forEach {
      if (limit < it.size) {
        limit = it.size
      }
    }

    fout.write("<table>\n".toByteArray())
    data.forEach {
      fout.write(" <tr>\n".toByteArray())
      for (i in 0..limit) {
        var cell = ""
        if (i < it.size) {
          cell = it[i]
        }
        fout.write("  <td>$cell</td>\n".toByteArray())
      }
      fout.write(" </tr>\n".toByteArray())
    }
    fout.write("</table>\n".toByteArray())
  }

  override fun header(title: String) {
  }

}