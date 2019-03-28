import org.junit.Assert.assertThat
import org.junit.Test
import java.io.PrintStream
import org.junit.Before
import java.io.ByteArrayOutputStream
import java.io.InputStream
import org.hamcrest.CoreMatchers.`is` as Is


class ExcelQueryTest {
  private val outContent = ByteArrayOutputStream()
  private val errContent = ByteArrayOutputStream()

  private var defaultSysin: InputStream? = null
  private var defaultSysout: PrintStream? = null

  @Before
  fun setup() {
    defaultSysin = System.`in`
    defaultSysout = System.out

    System.setOut(PrintStream(outContent))
    System.setErr(PrintStream(errContent))
  }
  @Test
  fun test1() {
    ExcelQuery.main(arrayOf("select A from `./data/sample1.xlsx`.Tokyo"))

    var result = outContent.toString().replace("\r\n", "\n")
    assertThat(result, Is("A\n1\n2\n3\n"))
  }
}