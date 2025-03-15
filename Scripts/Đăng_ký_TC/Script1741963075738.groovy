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
def saveTestResultToExcel(String email, String username, String phone, String password, boolean testResult, String message) {
	String projectPath = System.getProperty('user.dir')
	String excelFilePath = projectPath + '/Reports/TestResults.xlsx'
	File excelFile = new File(excelFilePath)

	Workbook workbook
	Sheet sheet

	if (excelFile.exists()) {
		FileInputStream fis = new FileInputStream(excelFile)
		workbook = new XSSFWorkbook(fis)
		sheet = workbook.getSheet('Register Test Results') // Sửa tên sheet
		fis.close()
	} else {
		workbook = new XSSFWorkbook()
		sheet = workbook.createSheet('Register Test Results')

		Row headerRow = sheet.createRow(0)
		headerRow.createCell(0).setCellValue('STT')
		headerRow.createCell(1).setCellValue('Thời gian')
		headerRow.createCell(2).setCellValue('Email')
		headerRow.createCell(3).setCellValue('Username')
		headerRow.createCell(4).setCellValue('Phone')
		headerRow.createCell(5).setCellValue('Password')
		headerRow.createCell(6).setCellValue('Kết quả')
		headerRow.createCell(7).setCellValue('Ghi chú')
		headerRow.createCell(8).setCellValue('Alert Message')
	}

	int lastRowNum = sheet.getLastRowNum()
	Row newRow = sheet.createRow(lastRowNum + 1)
	Date now = new Date()
	SimpleDateFormat dateFormat = new SimpleDateFormat('dd/MM/yyyy HH:mm:ss')
	String currentTime = dateFormat.format(now)

	newRow.createCell(0).setCellValue(lastRowNum + 1)
	newRow.createCell(1).setCellValue(currentTime)
	newRow.createCell(2).setCellValue(email)
	newRow.createCell(3).setCellValue(username)
	newRow.createCell(4).setCellValue(phone)
	newRow.createCell(5).setCellValue(password)

	CellStyle successStyle = workbook.createCellStyle()
	successStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex())
	successStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND)

	CellStyle failStyle = workbook.createCellStyle()
	failStyle.setFillForegroundColor(IndexedColors.RED.getIndex())
	failStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND)

	Cell resultCell = newRow.createCell(6)
	if (testResult) {
		resultCell.setCellValue('PASS')
		resultCell.setCellStyle(successStyle)
	} else {
		resultCell.setCellValue('FAIL')
		resultCell.setCellStyle(failStyle)
	}

	newRow.createCell(7).setCellValue(message)
	// Thêm nội dung alert nếu có
	  if (message.contains('Alert:')) {
		  newRow.createCell(8).setCellValue(message.substring(message.indexOf('Alert:') + 7))
	  }

	for (int i = 0; i < 9; i++) {
		sheet.autoSizeColumn(i)
	}

	FileOutputStream fos = new FileOutputStream(excelFile)
	workbook.write(fos)
	workbook.close()
	fos.close()

	println('Đã lưu kết quả test vào file: ' + excelFilePath)
}

// Hàm kiểm tra định dạng email (sử dụng regex)
def isValidEmail(String email) {
	String emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/
	return email.matches(emailRegex)
}
// Xử lí các alert, bạn cũng có thể bỏ hàm này
def handleAlert() {
	try {
		if (WebUI.waitForAlert(5)) {
			String alertText = WebUI.getAlertText()
			WebUI.acceptAlert()
			return alertText
		}
		return null
	} catch (Exception e) {
		return null
	}
}
// Hàm kiểm tra đăng ký thành công/thất bại (SỬ DỤNG OBJECT REPOSITORY TỪ RECORDER)
def checkRegisterSuccess() {
	// Kiểm tra xem có thông báo lỗi chung không (sau khi submit)
	if (WebUI.verifyElementPresent(findTestObject('Page_Register/p_ErrorMessage_Recorder'), 5, FailureHandling.OPTIONAL)) { // Thay bằng Object Repository của bạn
		String errorMessage = WebUI.getText(findTestObject('Page_Register/p_ErrorMessage_Recorder'))  // Thay bằng Object Repository của bạn
		if (errorMessage.contains('An error occurred. Please try again') || errorMessage.contains('Error: Email is already registered!') || errorMessage.contains('Error: Username is already registered!')) {
			return [false, errorMessage]
		}
	}

	 // Kiểm tra xem sau khi đăng ký có chuyển hướng đến trang đăng nhập không (HOẶC PHẦN TỬ KHÁC)
	 WebUI.delay(3)
	 if (WebUI.verifyElementPresent(findTestObject('Object Repository/Page_Rolex/button_ng Nhp'), 5, FailureHandling.OPTIONAL)){// Object Repository của bạn
		 return [true, "Đăng Ký thành công"]
	 }
		 // Nếu không có lỗi nào ở trên, kiểm tra alert (nếu có)
   String alertText = handleAlert()

	if (alertText != null) {
		if (alertText.toLowerCase().contains('success') || alertText.toLowerCase().contains('thành công')) {
			return [true, "Alert: " + alertText]
		} else if (alertText.toLowerCase().contains('fail') || alertText.toLowerCase().contains('thất bại')) {
			return [false, "Alert: " + alertText]
		}
	}
	return [false, "Đăng ký thất bại (không rõ nguyên nhân)"]
}

// Code test đăng ký (sử dụng Data File)
try {
	WebUI.openBrowser('')
	WebUI.navigateToUrl('http://127.0.0.1:5173/')
	WebUI.click(findTestObject('Object Repository/Page_Rolex/i_Lin H_fa-regular fa-user')) // Object Repository của bạn
	WebUI.click(findTestObject('Object Repository/Page_Rolex/span_ng K')) // Object Repository của bạn

	// Sử dụng biến từ Data File (đã được liên kết trong tab Variables)
	WebUI.setText(findTestObject('Object Repository/Page_Rolex/input_Email_email'), email) // Object Repository của bạn
	WebUI.setText(findTestObject('Object Repository/Page_Rolex/input_Tn Ngi Dng_username'), username) // Object Repository của bạn
	WebUI.setText(findTestObject('Object Repository/Page_Rolex/input_S in Thoi_phone'), phone) // Object Repository của bạn
	WebUI.setText(findTestObject('Object Repository/Page_Rolex/input_Mt Khu_password'), password) // Object Repository của bạn

	// 1. Kiểm tra định dạng email *trước*
	if (!isValidEmail(email)) {
		saveTestResultToExcel(email, username, phone, password, false, "Invalid email format")
		return // Kết thúc test case nếu email không hợp lệ
	}

	// 2. Kiểm tra bỏ trống trường *trước*
	boolean hasEmptyFieldError = false
	String emptyFieldMessage = ""

	if (WebUI.verifyElementHasAttribute(findTestObject('Object Repository/Page_Rolex/input_Email_email'), 'required', 3)) {
		if (!email) {
			hasEmptyFieldError = true
			emptyFieldMessage += "Email: Please fill out this field. "
		}
	}
	if (WebUI.verifyElementHasAttribute(findTestObject('Object Repository/Page_Rolex/input_Tn Ngi Dng_username'), 'required', 3)) {
		if (!username) {
			hasEmptyFieldError = true
			emptyFieldMessage += "Username: Please fill out this field. "
		}
	}
	if (WebUI.verifyElementHasAttribute(findTestObject('Object Repository/Page_Rolex/input_S in Thoi_phone'), 'required', 3)) {
		if (!phone) {
			hasEmptyFieldError = true
			emptyFieldMessage += "Phone: Please fill out this field. "
		}
	}
	if (WebUI.verifyElementHasAttribute(findTestObject('Object Repository/Page_Rolex/input_Mt Khu_password'), 'required', 3)) {
		if (!password) {
			hasEmptyFieldError = true
			emptyFieldMessage += "Password: Please fill out this field. "
		}
	}

	if (hasEmptyFieldError) {
		saveTestResultToExcel(email, username, phone, password, false, emptyFieldMessage.trim())
		println("Test case with empty fields. Data: email=" + email + ", username=" + username + ", phone=" + phone + ", password=" + password)
		return // Kết thúc test case nếu có trường bỏ trống
	}


	// Click nút đăng ký
	WebUI.click(findTestObject('Object Repository/Page_Rolex/button_ng K'))  // Object Repository của bạn

	// 3. Kiểm tra kết quả đăng ký (sau khi click)
	def registerResult = checkRegisterSuccess()
	boolean registerSuccess = registerResult[0]
	String message = registerResult[1]

	// Lưu kết quả vào Excel
	saveTestResultToExcel(email, username, phone, password, registerSuccess, message)

}
catch (Exception e) {
	// Xử lý ngoại lệ (lấy giá trị từ TestCaseContext nếu có, hoặc để rỗng)
	String email = ""
	String username = ""
	String phone = ""
	String password = ""

	 try{ //Cố gắng lấy giá trị từ context, nếu không có, exception sẽ được throw và xử lý ở catch bên ngoài
		 email =  TestCaseContext.getVariableValue("email")
		 username = TestCaseContext.getVariableValue("username")
		 phone =  TestCaseContext.getVariableValue("phone")
		 password = TestCaseContext.getVariableValue("password")
	 }
	 catch(Exception ex){
		 //Không làm gì cả, giữ giá trị rỗng
		 println("Could not get variables from TestCaseContext: " + ex.getMessage());
	 }
	saveTestResultToExcel(email, username, phone, password, false, 'Lỗi: ' + e.getMessage())
}
finally {
	WebUI.closeBrowser()
}