
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileInputStream
import java.io.FileOutputStream

/*
fun main(args: Array<String>) {
    val xlWb = XSSFWorkbook()
    val xlWs = xlWb.createSheet()
        xlWs.createRow(0).createCell(0).setCellValue("Привет жабам!")
    val output = FileOutputStream("./test.xlsx")
    xlWb.write(output)
    xlWb.close()
}
 */
fun main(args: Array<String>) {
    val input = FileInputStream("./test.xlsx")
    val xlWb = WorkbookFactory.create(input)
    val xlWs = xlWb.getSheetAt(0)
    println(xlWs.getRow(0).getCell(0))
}