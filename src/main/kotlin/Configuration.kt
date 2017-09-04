class Configuration(args: Array<String>) {
  var divideItems = 0
  var inputFile = ""
  var outputDirectory = ""
  var sheets: List<String> = listOf()

  init {
    var i = 0
    while(i < args.size) {
      var arg = args[i]
      if (arg.startsWith("-")) {
        arg = arg.substring(1, arg.length)
        when(arg) {
          "sheets" -> {
            sheets = args[i + 1].split(",")
            i++
          }
          "in" -> {
            inputFile = args[i + 1]
            i++
          }
          "out" -> {
            outputDirectory = args[i + 1]
            i++
          }
          "divide" -> {
            divideItems = args[i + 1].toInt()
            i++
          }
        }
      }
      i++
    }
  }
}