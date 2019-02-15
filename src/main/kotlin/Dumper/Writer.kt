package Dumper

interface Writer: java.io.Closeable {
  fun header(title:String)
  fun write(s: List<List<String>>)
  fun sheetTitle(title: String)
}