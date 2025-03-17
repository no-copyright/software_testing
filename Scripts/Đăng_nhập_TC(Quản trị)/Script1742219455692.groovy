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
import java.io.File

/**
 * Hàm lưu kết quả test vào file Excel cho kịch bản "Đăng nhập Admin"
 */
def saveAdminLoginResultToExcel(String username, String password, boolean testResult, String message) {
    // Định nghĩa đường dẫn đến file Excel
    String projectPath = System.getProperty('user.dir')
    String excelFilePath = projectPath + '/Reports/TestResults.xlsx'
    File excelFile = new File(excelFilePath)

    Workbook workbook
    Sheet sheet

    // Kiểm tra xem file Excel đã tồn tại chưa
    if (excelFile.exists()) {
        // Nếu file đã tồn tại, mở file
        FileInputStream fis = new FileInputStream(excelFile)
        workbook = new XSSFWorkbook(fis)
        // Tạo sheet "Admin Login Test Results" nếu chưa có
        if (workbook.getSheet("Admin Login Test Results") == null) {
            sheet = workbook.createSheet('Admin Login Test Results')
        } else {
            sheet = workbook.getSheet('Admin Login Test Results')
        }
        fis.close()
    } else {
        // Nếu file chưa tồn tại, tạo file mới
        workbook = new XSSFWorkbook()
        sheet = workbook.createSheet('Admin Login Test Results')

        // Tạo header cho file Excel
        Row headerRow = sheet.createRow(0)
        headerRow.createCell(0).setCellValue('STT')
        headerRow.createCell(1).setCellValue('Thời gian')
        headerRow.createCell(2).setCellValue('Username')
        headerRow.createCell(3).setCellValue('Password')
        headerRow.createCell(4).setCellValue('Kết quả')
        headerRow.createCell(5).setCellValue('Ghi chú')
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

    println('Đã lưu kết quả test vào file: ' + excelFilePath)
}

/**
 * Kịch bản kiểm thử "Đăng nhập Admin"
 */
try {
    WebUI.openBrowser('')
	WebUI.maximizeWindow()
    WebUI.navigateToUrl('http://localhost:5174/')

    WebUI.setText(findTestObject('Object Repository/Page_Rolex Admin/input_Tn ng Nhp_username'), username)
    WebUI.setText(findTestObject('Object Repository/Page_Rolex Admin/input_Mt Khu_password'), password)
    WebUI.click(findTestObject('Object Repository/Page_Rolex Admin/button_ng Nhp'))

    // **Kiểm tra đăng nhập thành công hay thất bại**
    boolean loginSuccess = false
    String message = ""

    // Kiểm tra xem có thông báo lỗi "Please fill out this field" hay không
    if (WebUI.verifyElementPresent(findTestObject('Object Repository/Page_Rolex Admin/div_Please fill out this field'), 5, FailureHandling.OPTIONAL)) {
        message = "Đăng nhập thất bại (Chưa nhập đủ thông tin)"
        loginSuccess = false
    }
    // Kiểm tra đăng nhập thất bại:
    else if (WebUI.verifyElementPresent(findTestObject('Page_Rolex Admin/check_login'), 5, FailureHandling.OPTIONAL)) {
        WebUI.check(findTestObject('Page_Rolex Admin/check_login')) // giữ nguyên sự kiện check vì bạn muốn thế
        message = "Đăng nhập thất bại (Thông tin sai)"
        loginSuccess = false
    } else {
        // Nếu không có check_login, thử click vào ảnh để xác định đăng nhập thành công
        try {
            WebUI.click(findTestObject('Page_Rolex Admin/img'))
            message = "Đăng nhập thành công (Click được vào ảnh)"
            loginSuccess = true
        } catch (Exception imgError) {
            message = "Đăng nhập thất bại (Không click được vào ảnh): " + imgError.getMessage()
            loginSuccess = false
        }
    }

    // Lưu kết quả test vào Excel
    saveAdminLoginResultToExcel(username, password, loginSuccess, message)

} catch (Exception e) {
    // Nếu có lỗi xảy ra trong quá trình test
    saveAdminLoginResultToExcel(username, password, false, 'Lỗi: ' + e.getMessage())
} finally {
    // Đóng trình duyệt sau khi test
    WebUI.closeBrowser()
}