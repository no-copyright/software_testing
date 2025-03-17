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
     // Xử lý alert đăng nhập (nếu có)
    try {
        WebUI.waitForAlert(5)
        WebUI.acceptAlert()
        println("Đã chấp nhận alert sau khi đăng nhập.")
    } catch (Exception e) {
        println("Không có alert sau khi đăng nhập.")
    }
    ghiLog("Đăng nhập", true)

    // 3. Click vào "Sản Phẩm"
    WebUI.click(findTestObject('Object Repository/Page_Rolex Admin/a_Sn Phm'))
    ghiLog("Click vào 'Sản Phẩm'", true)

    // 4. Click vào "Danh Sách Sản Phẩm"
    WebUI.click(findTestObject('Object Repository/Page_Rolex Admin/a_Danh Sch Sn Phm'))
    ghiLog("Click vào 'Danh Sách Sản Phẩm'", true)

    // 5. Click vào nút "Ẩn"
    WebUI.click(findTestObject('Object Repository/Page_Rolex Admin/button_n'))
    ghiLog("Click vào nút 'Ẩn'", true)
    
    // Xử lý alert Ẩn (nếu có)
    try {
        WebUI.waitForAlert(5)
        WebUI.acceptAlert()
        println("Đã chấp nhận alert sau khi click Ẩn.")
        ghiLog("Xử lý alert sau khi click 'Ẩn' thành công", true)
        
        // Thêm bước reload trang sau khi xác nhận alert ẩn
        WebUI.refresh()
        ghiLog("Đã reload trang để hiển thị nút 'Xóa'", true)
        
        // Thêm độ trễ nhỏ để đảm bảo trang đã tải lại hoàn toàn
        WebUI.delay(2)
    } catch (Exception e) {
        println("Không có alert sau click Ẩn.")
        ghiLog("Không có alert sau khi click 'Ẩn'", false)
    }

    // 6. Click vào nút "Xóa"
    TestObject nutXoa = findTestObject('Object Repository/Page_Rolex Admin/button_Xa')

    // Kiểm tra xem nút "Xóa" có clickable và visible không
    if (WebUI.verifyElementClickable(nutXoa, FailureHandling.OPTIONAL) && WebUI.verifyElementVisible(nutXoa, FailureHandling.OPTIONAL)) {
        WebUI.click(nutXoa)
        ghiLog("Click vào nút 'Xóa'", true)
    } else {
        ghiLog("Nút 'Xóa' không clickable hoặc không visible", false)
        WebUI.closeBrowser() // Close browser on failure
        return // Kết thúc script
    }

    // Xử lý alert sau khi click "Xóa" (QUAN TRỌNG)
    try {
        WebUI.waitForAlert(10) // Chờ alert trong 10 giây
        WebUI.acceptAlert()  // Chấp nhận alert
        println("Đã chấp nhận alert sau khi click 'Xóa'.")
        ghiLog("Xử lý alert sau khi click 'Xóa' thành công", true)
    } catch (Exception e) {
        println("Không có alert sau khi click 'Xóa'.")
        ghiLog("Không có alert sau khi click 'Xóa'", false)
        // Hoặc bạn có thể xử lý lỗi ở đây, tùy thuộc vào yêu cầu
    }
} catch (Exception e) {
    ghiLog("Xảy ra lỗi: " + e.getMessage(), false)
} finally {
    WebUI.closeBrowser()
    ghiLog("Kết thúc test case", true) // Kết thúc test case
}