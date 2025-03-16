import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject
import com.kms.katalon.core.checkpoint.Checkpoint as Checkpoint
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling as FailureHandling
import com.kms.katalon.core.testcase.TestCase as TestCase
import com.kms.katalon.core.testdata.TestData as TestData
import com.kms.katalon.core.testng.keyword.TestNGBuiltinKeywords as TestNGKW
import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows
import internal.GlobalVariable as GlobalVariable
import org.openqa.selenium.Keys as Keys

// Import các thư viện cần thiết để làm việc với Excel
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.*
import java.io.FileInputStream
import java.io.FileOutputStream
import java.text.SimpleDateFormat
import java.util.Date

// Hàm để lưu kết quả test vào file Excel
def saveTestResultToExcel(String username, String password, boolean testResult, String message) {
    // Định nghĩa đường dẫn đến file Excel
    String projectPath = System.getProperty('user.dir')
    String excelFilePath = projectPath + '/Reports/TestResults.xlsx'
    File excelFile = new File(excelFilePath)
    String sheetName = "Test Results"

    Workbook workbook
    Sheet sheet

    // Kiểm tra xem file Excel đã tồn tại chưa
    if (excelFile.exists()) {
        // Nếu file đã tồn tại, mở file
        FileInputStream fis = new FileInputStream(excelFile)
        workbook = new XSSFWorkbook(fis)
        sheet = workbook.getSheet(sheetName)
        if (sheet == null) {
            sheet = workbook.createSheet(sheetName)
            createHeaderRow(sheet) // Create header if sheet is new
        }
        fis.close()
    } else {
        // Nếu file chưa tồn tại, tạo file mới
        workbook = new XSSFWorkbook()
        sheet = workbook.createSheet(sheetName)
        createHeaderRow(sheet)  //Create header row
    }

    // Lấy số hàng hiện tại của sheet
    int lastRowNum = sheet.getLastRowNum()

    // Tạo hàng mới để lưu kết quả test
    Row newRow = sheet.createRow(lastRowNum + 1)

    // Lấy thời gian hiện tại
    Date now = new Date()
    SimpleDateFormat dateFormat = new SimpleDateFormat('dd/MM/yyyy HH:mm:ss')
    String currentTime = dateFormat.format(now)

    // Điền thông tin vào các ô
    newRow.createCell(0).setCellValue(lastRowNum + 1) // STT
    newRow.createCell(1).setCellValue(currentTime) // Thời gian test
    newRow.createCell(2).setCellValue(username) // Username
    newRow.createCell(3).setCellValue(password) // Password

    // Cell style cho kết quả test
    CellStyle successStyle = workbook.createCellStyle()
    successStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex())
    successStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND)

    CellStyle failStyle = workbook.createCellStyle()
    failStyle.setFillForegroundColor(IndexedColors.RED.getIndex())
    failStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND)

    // Điền kết quả test
    Cell resultCell = newRow.createCell(4)
    if (testResult) {
        resultCell.setCellValue('PASS')
        resultCell.setCellStyle(successStyle)
    } else {
        resultCell.setCellValue('FAIL')
        resultCell.setCellStyle(failStyle)
    }

    // Điền ghi chú
    newRow.createCell(5).setCellValue(message)

    // Tự động điều chỉnh độ rộng của các cột
    for (int i = 0; i < 6; i++) {
        sheet.autoSizeColumn(i)
    }

    // Lưu file Excel
    FileOutputStream fos = new FileOutputStream(excelFile)
    workbook.write(fos)
    workbook.close()
    fos.close()

    println('Đã lưu kết quả test vào file: ' + excelFilePath + " vào sheet " + sheetName)
}

//Helper function to create header row
def createHeaderRow(Sheet sheet) {
    Row headerRow = sheet.createRow(0)
    headerRow.createCell(0).setCellValue('STT')
    headerRow.createCell(1).setCellValue('Thời gian')
    headerRow.createCell(2).setCellValue('Username')
    headerRow.createCell(3).setCellValue('Password')
    headerRow.createCell(4).setCellValue('Kết quả')
    headerRow.createCell(5).setCellValue('Ghi chú')
}

try {
    WebUI.openBrowser('')
    WebUI.navigateToUrl('http://localhost:5173/')

    WebUI.click(findTestObject('Object Repository/Page_Rolex/i_Lin H_fa-regular fa-user'))

    WebUI.setText(findTestObject('Object Repository/Page_Rolex/input_Tn Ngi Dng_username'), username)

    WebUI.setText(findTestObject('Object Repository/Page_Rolex/input_Mt Khu_password'), password)

    WebUI.click(findTestObject('Object Repository/Page_Rolex/button_ng Nhp'))

    // Handle alert sau khi đăng nhập (nếu có)
    try {
        WebUI.waitForAlert(5)
        WebUI.acceptAlert()
        println("Đã chấp nhận alert sau khi đăng nhập.")
    } catch (Exception e) {
        println("Không có alert sau khi đăng nhập.")
    }

    WebUI.click(findTestObject('Object Repository/Page_Rolex/a_Sn Phm'))

    WebUI.click(findTestObject('Object Repository/Page_Rolex/button_Thm Vo Gi'))

    WebUI.click(findTestObject('Object Repository/Page_Rolex/button_Thm Vo Gi_1'))

    WebUI.click(findTestObject('Object Repository/Page_Rolex/button_t Hng'))

    // Handle alert sau khi đặt hàng
    try {
        WebUI.waitForAlert(5)
        WebUI.acceptAlert()
        println("Đã chấp nhận alert sau khi đặt hàng.")
    } catch (Exception e) {
        println("Không có alert sau khi đặt hàng.")
    }

    WebUI.click(findTestObject('Object Repository/Page_Rolex/i_Lin H_fa-solid fa-cart-shopping'))

    WebUI.click(findTestObject('Object Repository/Page_Rolex/button_Thanh Ton'))

    WebUI.switchToWindowTitle('Tạo mới đơn hàng')

    WebUI.click(findTestObject('Object Repository/Page_To mi n hng/input_Chn Phng thc thanh ton_bankCode'))

    WebUI.click(findTestObject('Object Repository/Page_To mi n hng/button_Thanh ton'))

    WebUI.click(findTestObject('Object Repository/Page_Chn phng thc thanh ton (Test)/div_Th ni a v ti khon ngn hng_list-bank-item-inner'))

    WebUI.setText(findTestObject('Object Repository/Page_Thanh ton qua Ngn hng NCB/input_Th ni a_card_number_mask'), card_number)

    WebUI.setText(findTestObject('Object Repository/Page_Thanh ton qua Ngn hng NCB/input_S th_cardHolder'), cardholder_name)

    WebUI.setText(findTestObject('Object Repository/Page_Thanh ton qua Ngn hng NCB/input_Tn ch th_cardDate'), release_date)

    WebUI.click(findTestObject('Object Repository/Page_Thanh ton qua Ngn hng NCB/span_Tip tc'))

    WebUI.click(findTestObject('Object Repository/Page_Thanh ton qua Ngn hng NCB/span_ng   Tip tc'))

    // Chờ 7 giây để kiểm tra xem có chuyển qua màn hình OTP không
    TestObject otpInput = findTestObject('Object Repository/Page_Xc thc OTP/input_Xc thc OTP_otpvalue')
    boolean isOtpPageDisplayed = WebUI.waitForElementPresent(otpInput, 7)

    // Nếu không chuyển được qua trang OTP, chứng tỏ thông tin thẻ không hợp lệ
    if (!isOtpPageDisplayed) {
        // Kiểm tra có hiển thị thông báo lỗi không
        TestObject errorMessageObject = findTestObject('Object Repository/Page_Thanh ton qua Ngn hng NCB/div_Error_Message_Locator') // Thay thế bằng locator thực tế
        boolean errorMessagePresent = WebUI.waitForElementPresent(errorMessageObject, 3)
        
        String errorMessage = "Không thể chuyển qua trang nhập OTP sau khi nhập thông tin thẻ"
        if (errorMessagePresent) {
            errorMessage = WebUI.getText(errorMessageObject)
        }
        
        saveTestResultToExcel(username, password, false, 'Thông tin thẻ không hợp lệ: ' + errorMessage)
        WebUI.closeBrowser()
        return
    }

    // Nếu đã chuyển sang trang OTP, tiếp tục nhập OTP
    WebUI.setText(otpInput, otp)

    WebUI.click(findTestObject('Object Repository/Page_Xc thc OTP/span_Thanh ton'))

    // Kiểm tra xem thanh toán thành công hay không
    try {
        WebUI.waitForElementPresent(findTestObject('Object Repository/Page_/p_GD thnh cng'), 5)
        WebUI.click(findTestObject('Object Repository/Page_/p_GD thnh cng'))

        WebUI.switchToWindowTitle('Rolex')

        WebUI.click(findTestObject('Object Repository/Page_Rolex/button_'))

        WebUI.click(findTestObject('Object Repository/Page_Rolex/i_Lin H_fa-solid fa-cart-shopping'))

        WebUI.waitForElementPresent(findTestObject('Object Repository/Page_Rolex/p_Trng Thi  hon thnh'), 5)
        WebUI.click(findTestObject('Object Repository/Page_Rolex/p_Trng Thi  hon thnh'))

        // Lưu kết quả thanh toán thành công
        saveTestResultToExcel(username, password, true, 'Thanh toán thành công, đơn hàng hoàn thành')
    } catch (Exception e) {
        // Lưu kết quả thanh toán thất bại
        saveTestResultToExcel(username, password, false, 'Thanh toán thất bại: ' + e.getMessage())
    }

} catch (Exception e) {
    // Nếu có lỗi xảy ra trong quá trình test
    saveTestResultToExcel(username, password, false, 'Lỗi: ' + e.getMessage())
} finally {
    // Đóng trình duyệt sau khi test
    WebUI.closeBrowser()
}