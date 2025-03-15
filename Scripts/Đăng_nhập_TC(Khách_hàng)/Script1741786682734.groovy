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
    
    Workbook workbook
    Sheet sheet
    // Test
    // Kiểm tra xem file Excel đã tồn tại chưa
    if (excelFile.exists()) {
        // Nếu file đã tồn tại, mở file
        FileInputStream fis = new FileInputStream(excelFile)
        workbook = new XSSFWorkbook(fis)
        sheet = workbook.getSheet('Login Test Results')
        fis.close()
    } else {
        // Nếu file chưa tồn tại, tạo file mới
        workbook = new XSSFWorkbook()
        sheet = workbook.createSheet('Login Test Results')
        
        // Tạo header cho file Excel
        Row headerRow = sheet.createRow(0)
        headerRow.createCell(0).setCellValue('STT')
        headerRow.createCell(1).setCellValue('Thời gian')
        headerRow.createCell(2).setCellValue('Username')
        headerRow.createCell(3).setCellValue('Password')
        headerRow.createCell(4).setCellValue('Kết quả')
        headerRow.createCell(5).setCellValue('Ghi chú')
        headerRow.createCell(6).setCellValue('Alert Message')
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
    
    // Thêm nội dung alert nếu có
    if (message.contains('Alert:')) {
        newRow.createCell(6).setCellValue(message.substring(message.indexOf('Alert:') + 7))
    }
    
    // Tự động điều chỉnh độ rộng của các cột
    for (int i = 0; i < 7; i++) {
        sheet.autoSizeColumn(i)
    }
    
    // Lưu file Excel
    FileOutputStream fos = new FileOutputStream(excelFile)
    workbook.write(fos)
    workbook.close()
    fos.close()
    
    println('Đã lưu kết quả test vào file: ' + excelFilePath)
}

// Hàm kiểm tra alert và trả về nội dung alert
def handleAlert() {
    try {
        // Chờ alert xuất hiện với timeout 5 giây
        if (WebUI.waitForAlert(5)) {
            // Lấy nội dung alert
            String alertText = WebUI.getAlertText()
            
            // Chấp nhận alert (bấm OK)
            WebUI.acceptAlert()
            
            return alertText
        }
        return null
    } catch (Exception e) {
        return null
    }
}

// Hàm kiểm tra đăng nhập thành công hay thất bại
def checkLoginSuccess() {
    // Xử lý alert nếu có
    String alertText = handleAlert()
    
    // Kiểm tra nội dung alert
    if (alertText != null) {
        if (alertText.toLowerCase().contains('success') || alertText.toLowerCase().contains('thành công')) {
            // Nếu alert chứa từ "success" hoặc "thành công", coi như đăng nhập thành công
            return [true, "Alert: " + alertText]
        } else if (alertText.toLowerCase().contains('fail') || alertText.toLowerCase().contains('thất bại')) {
            // Nếu alert chứa từ "fail" hoặc "thất bại", coi như đăng nhập thất bại
            return [false, "Alert: " + alertText]
        }
    }
    
    // Nếu không có alert hoặc alert không chứa từ khóa nhận diện, kiểm tra bằng phần tử trên trang
    try {
        // Đợi 3 giây để trang load sau khi đăng nhập
        WebUI.delay(3)
        
        // Kiểm tra đăng nhập thành công bằng cách tìm một phần tử chỉ xuất hiện sau khi đăng nhập
        // Ví dụ: Có thể là avatar người dùng, nút logout, tên người dùng, v.v.
        if (WebUI.verifyElementPresent(findTestObject('Object Repository/Page_Rolex/div_user_avatar'), 5, FailureHandling.OPTIONAL)) {
            return [true, alertText ? "Alert: " + alertText : "Đăng nhập thành công"]
        }
        return [false, alertText ? "Alert: " + alertText : "Đăng nhập thất bại"]
    } catch (Exception e) {
        return [false, "Lỗi kiểm tra đăng nhập: " + e.getMessage()]
    }
}

// Code test đăng nhập
try {
    WebUI.openBrowser('')
    WebUI.navigateToUrl('http://localhost:5173/')
    WebUI.click(findTestObject('Object Repository/Page_Rolex/i_Lin H_fa-regular fa-user'))
    WebUI.setText(findTestObject('Object Repository/Page_Rolex/input_Tn Ngi Dng_username'), username)
    WebUI.setText(findTestObject('Object Repository/Page_Rolex/input_Mt Khu_password'), password)
    WebUI.click(findTestObject('Object Repository/Page_Rolex/button_ng Nhp'))
    
    // Kiểm tra kết quả đăng nhập
    def loginResult = checkLoginSuccess()
    boolean loginSuccess = loginResult[0]
    String message = loginResult[1]
    
    // Lưu kết quả test vào Excel
    saveTestResultToExcel(username, password, loginSuccess, message)
    
} catch (Exception e) {
    // Nếu có lỗi xảy ra trong quá trình test
    saveTestResultToExcel(username, password, false, 'Lỗi: ' + e.getMessage())
} finally {
    // Đóng trình duyệt sau khi test
    WebUI.closeBrowser()
}