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

import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.*
import java.io.FileInputStream
import java.io.FileOutputStream
import java.text.SimpleDateFormat
import java.util.Date
import java.io.File

// Hàm lưu kết quả test vào file Excel
def saveEditProductResultToExcel(String name, String image, String title, String brand, String price, String category, boolean testResult, String message) {
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
        // Tạo sheet "Edit Product Test Results" nếu chưa có
        if (workbook.getSheet("Edit Product Test Results") == null) {
            sheet = workbook.createSheet('Edit Product Test Results')
        } else {
            sheet = workbook.getSheet('Edit Product Test Results')
        }
        fis.close()
    } else {
        // Nếu file chưa tồn tại, tạo file mới
        workbook = new XSSFWorkbook()
        sheet = workbook.createSheet('Edit Product Test Results')
        
        // Tạo header cho file Excel
        Row headerRow = sheet.createRow(0)
        headerRow.createCell(0).setCellValue('STT')
        headerRow.createCell(1).setCellValue('Thời gian')
        headerRow.createCell(2).setCellValue('Tên sản phẩm')
        headerRow.createCell(3).setCellValue('Hình ảnh')
        headerRow.createCell(4).setCellValue('Tiêu đề')
        headerRow.createCell(5).setCellValue('Nhãn hiệu')
        headerRow.createCell(6).setCellValue('Giá tiền')
        headerRow.createCell(7).setCellValue('Danh mục')
        headerRow.createCell(8).setCellValue('Kết quả')
        headerRow.createCell(9).setCellValue('Ghi chú')
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
    newRow.createCell(2).setCellValue(name) // Tên sản phẩm
    newRow.createCell(3).setCellValue(image) // Hình ảnh
    newRow.createCell(4).setCellValue(title) // Tiêu đề
    newRow.createCell(5).setCellValue(brand) // Nhãn hiệu
    newRow.createCell(6).setCellValue(price) // Giá tiền
    newRow.createCell(7).setCellValue(category) // Danh mục
    
    // Cell style cho kết quả test
    CellStyle successStyle = workbook.createCellStyle()
    successStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex())
    successStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND)
    
    CellStyle failStyle = workbook.createCellStyle()
    failStyle.setFillForegroundColor(IndexedColors.RED.getIndex())
    failStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND)
    
    // Điền kết quả test
    Cell resultCell = newRow.createCell(8)
    if (testResult) {
        resultCell.setCellValue('PASS')
        resultCell.setCellStyle(successStyle)
    } else {
        resultCell.setCellValue('FAIL')
        resultCell.setCellStyle(failStyle)
    }
    
    newRow.createCell(9).setCellValue(message)

    // Tự động điều chỉnh độ rộng của các cột
    for (int i = 0; i < 10; i++) {
        sheet.autoSizeColumn(i)
    }

    // Lưu file Excel
    FileOutputStream fos = new FileOutputStream(excelFile)
    workbook.write(fos)
    workbook.close()
    fos.close()

    println('Đã lưu kết quả test vào file: ' + excelFilePath)
}

// Code test sửa sản phẩm
try {
    WebUI.openBrowser('')
	WebUI.maximizeWindow()
    WebUI.navigateToUrl('http://localhost:5174/')

    WebUI.setText(findTestObject('Object Repository/Page_Rolex Admin/input_Tn ng Nhp_username'), username)
    WebUI.setText(findTestObject('Object Repository/Page_Rolex Admin/input_Mt Khu_password'), password)
    WebUI.click(findTestObject('Object Repository/Page_Rolex Admin/button_ng Nhp'))
    WebUI.click(findTestObject('Object Repository/Page_Rolex Admin/i_Sn Phm_fa-solid fa-chevron-down'))
    WebUI.click(findTestObject('Object Repository/Page_Rolex Admin/a_Danh Sch Sn Phm'))
    WebUI.click(findTestObject('Object Repository/Page_Rolex Admin/a_Sa'))

    WebUI.setText(findTestObject('Object Repository/Page_Rolex Admin/input_Tn Sn Phm_product'), name)
    WebUI.uploadFile(findTestObject('Object Repository/Page_Rolex Admin/input_file (sua_san_pham)'), image)
    WebUI.setText(findTestObject('Object Repository/Page_Rolex Admin/input_Tiu_title'), title)
    WebUI.setText(findTestObject('Object Repository/Page_Rolex Admin/input_Nhn Hiu_brand'), brand)
    WebUI.setText(findTestObject('Object Repository/Page_Rolex Admin/input_Gi Tin_price'), price)
    WebUI.setText(findTestObject('Object Repository/Page_Rolex Admin/input_Danh Mc_category'), category)

    WebUI.click(findTestObject('Object Repository/Page_Rolex Admin/button_Cp Nhp'))

    // Kiểm tra xem Alert có xuất hiện hay không
    boolean isAlertPresent = WebUI.waitForAlert(5)

    if (isAlertPresent) {
        // Lấy text từ Alert (nếu muốn)
        String alertText = WebUI.getAlertText()
        WebUI.acceptAlert() // Chấp nhận Alert

        // Coi như sửa sản phẩm thành công nếu Alert xuất hiện
        saveEditProductResultToExcel(name, image, title, brand, price, category, true, "Sửa sản phẩm thành công (Alert: " + alertText + ")")
    } else {
        // Nếu không có Alert, coi như sửa sản phẩm thất bại
        saveEditProductResultToExcel(name, image, title, brand, price, category, false, "Sửa sản phẩm thất bại (Không có Alert)")
    }

} catch (Exception e) {
    // Kiểm tra nếu lỗi là do upload file không thành công
    if (e.getMessage().contains("Unable to upload file")) {
        saveEditProductResultToExcel(name, image, title, brand, price, category, false, "Sửa sản phẩm thất bại (Không thể tải file)")
    } else {
        // Nếu không phải lỗi upload file, ghi lại lỗi gốc
        saveEditProductResultToExcel(name, image, title, brand, price, category, false, 'Lỗi: ' + e.getMessage())
    }
} finally {
    // Đóng trình duyệt sau khi test
    WebUI.closeBrowser()
}
