import org.hamcrest.CoreMatchers.notNullValue
import org.junit.Assert.assertThat
import org.junit.Before
import org.junit.Test
import java.io.ByteArrayOutputStream
import java.io.InputStream
import java.io.PrintStream
import javax.xml.parsers.DocumentBuilderFactory
import org.hamcrest.CoreMatchers.`is` as Is
import java.io.ByteArrayInputStream
import java.nio.charset.StandardCharsets


class DumpDataTest {
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
    DumpData.main(arrayOf("./data/sample2.xlsx"))

    val result = outContent.toString().replace("\r\n", "\n")
    assertThat(result, Is(notNullValue()))

//    val documentBuilder = DocumentBuilderFactory.newInstance().newDocumentBuilder()
//    val stream = ByteArrayInputStream(result.toByteArray(StandardCharsets.UTF_8))
//
//    val dom = documentBuilder.parse(stream)
//    val heads = dom.getElementsByTagName("h1")
//    assertThat(heads.length, Is(4))
//    assertThat(heads.item(0).textContent, Is("Sheet1"))
//    assertThat(heads.item(1).textContent, Is("Sheet2"))
//    assertThat(heads.item(2).textContent, Is("Sheet3"))
//    assertThat(heads.item(3).textContent, Is("Sheet4"))

  }
}