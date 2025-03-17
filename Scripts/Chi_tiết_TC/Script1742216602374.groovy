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

// Hàm hỗ trợ để ghi log kết quả bằng tiếng Việt
def ghiLog(String thongBao, boolean ketQua) {
    String trangThai = ketQua ? "THÀNH CÔNG" : "THẤT BẠI"
    println(">>> ${trangThai}: ${thongBao}")
}

try {
    // 1. Mở trình duyệt và truy cập trang web
    WebUI.openBrowser('')
    WebUI.navigateToUrl('http://127.0.0.1:5174/')
    ghiLog("Mở trình duyệt và truy cập URL", true)

    // 2. Đăng nhập
    WebUI.setText(findTestObject('Object Repository/Page_Rolex Admin/input_Tn ng Nhp_username'), 'admin')
    WebUI.setEncryptedText(findTestObject('Object Repository/Page_Rolex Admin/input_Mt Khu_password'), 'yHe9suVz1Dkg/2HJIvR/uA==')
    WebUI.click(findTestObject('Object Repository/Page_Rolex Admin/button_ng Nhp'))
    // Xử lý alert đăng nhập (nếu có) - Nên để ở đây
    try {
        WebUI.waitForAlert(5)
        WebUI.acceptAlert()
        println("Đã chấp nhận alert sau khi đăng nhập.")
    } catch (Exception e) {
        println("Không có alert sau khi đăng nhập.")
    }
    ghiLog("Đăng nhập", true)

    // 3. Click vào "Chi Tiết"
    WebUI.click(findTestObject('Object Repository/Page_Rolex Admin/a_Chi Tit'))
    ghiLog("Click vào 'Chi Tiết'", true)

    // Chờ đợi trang chi tiết tải
    WebUI.delay(1) // Chờ 1 giây
    TestObject elementTrenTrangChiTiet = findTestObject('Object Repository/Page_Rolex Admin/element_tren_trang_chi_tiet') // THAY THẾ
    WebUI.waitForElementVisible(elementTrenTrangChiTiet, 10)

    // 4. Quay lại
    WebUI.back()
    ghiLog("Đã quay lại trang trước", true)

    // 5. Click vào "Hoàn Thành"
    WebUI.click(findTestObject('Object Repository/Page_Rolex Admin/button_Hon Thnh'))

    // Xử lý alert (QUAN TRỌNG: Dùng waitForAlert và xử lý cả trường hợp không có alert)
    try {
        WebUI.waitForAlert(10) // Chờ alert trong 10 giây
        WebUI.acceptAlert()
        println("Đã chấp nhận alert sau khi click 'Hoàn Thành'.")
        ghiLog("Xử lý alert sau khi click 'Hoàn Thành' thành công", true)
    } catch (Exception e) {
        println("Không có alert sau khi click 'Hoàn Thành'.") // Thông báo nếu không có alert
        ghiLog("Không có alert sau khi click 'Hoàn Thành'", false) // Ghi log là không có alert
    }

    // 6. Kiểm tra trạng thái (nếu cần, tùy thuộc vào ứng dụng)
    // ... (Phần này bạn có thể bỏ qua nếu alert đã xác nhận hành động)

} catch (Exception e) {
    ghiLog("Xảy ra lỗi: " + e.getMessage(), false)
} finally {
    WebUI.closeBrowser()
     ghiLog("Kết thúc test case", true)
}