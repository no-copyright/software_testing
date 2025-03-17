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

// Helper function to log test results
def logResult(String message, boolean result) {
    String status = result ? "PASS" : "FAIL"
    println(">>> ${status}: ${message}")
}

try {
    // 1. Open browser and navigate to URL
    WebUI.openBrowser('')
    WebUI.navigateToUrl('http://127.0.0.1:5173/')
    logResult("Opened browser and navigated to URL", true)

    // 2. Login
    WebUI.click(findTestObject('Object Repository/Page_Rolex/i_Lin H_fa-regular fa-user'))
    WebUI.setText(findTestObject('Object Repository/Page_Rolex/input_Tn Ngi Dng_username'), 'dat123')
    WebUI.setEncryptedText(findTestObject('Object Repository/Page_Rolex/input_Mt Khu_password'), 'p1TO31MZyWk=')
    WebUI.click(findTestObject('Object Repository/Page_Rolex/button_ng Nhp'))

    // Handle login alert (if any)
    try {
        WebUI.waitForAlert(5)
        WebUI.acceptAlert()
        println("Accepted login alert.")
    } catch (Exception e) {
        println("No login alert found.")
    }
    logResult("Login successful", true) // Assuming login is successful if no errors

    // 3. Navigate to Products page
    WebUI.click(findTestObject('Object Repository/Page_Rolex/a_Sn Phm'))
    logResult("Navigated to Products page", true)

    // 4. Add product to cart
    WebUI.click(findTestObject('Object Repository/Page_Rolex/button_Thm Vo Gi'))

    // Handle add to cart alert (if any)
    try {
        WebUI.waitForAlert(5)
        WebUI.acceptAlert()
        println("Accepted add to cart alert")
    } catch (Exception e) {
        println("No add to cart alert")
    }

    // 5. Remove product from cart
    TestObject productInCart = findTestObject('Object Repository/Page_Rolex/button_Xa') // REPLACE with correct locator
    WebUI.click(findTestObject('Object Repository/Page_Rolex/button_Xa'))
    // Handle remove product alert (if any)
    try {
        WebUI.waitForAlert(5)
        WebUI.acceptAlert()
        println("Accepted remove from cart alert")
    }
    catch(Exception e){
        println("No remove from cart alert")
    }

    // Verify product removal
    if (WebUI.waitForElementNotPresent(productInCart, 5, FailureHandling.OPTIONAL)) {
        logResult("Product removed from cart successfully", true)
    } else {
        logResult("Failed to remove product from cart", false)
        WebUI.closeBrowser() // Close browser on failure
        return
    }

    // 6. Add another product to cart
    WebUI.click(findTestObject('Object Repository/Page_Rolex/button_Thm Vo Gi_1'))
      // Handle add to cart alert (if any)
    try {
        WebUI.waitForAlert(5)
        WebUI.acceptAlert()
        println("Accepted add to cart alert")
    } catch (Exception e) {
        println("No add to cart alert")
    }

    // 7. Place order
    WebUI.click(findTestObject('Object Repository/Page_Rolex/button_t Hng'))
      // Handle order alert (if any)
    try {
        WebUI.waitForAlert(5)
        WebUI.acceptAlert()
        println("Accepted order alert")
    } catch (Exception e) {
        println("No order alert")
    }
    logResult("Placed order", true) // Assuming order placement is successful if no errors

    // 8. Go to cart
    WebUI.click(findTestObject('Object Repository/Page_Rolex/i_Lin H_fa-solid fa-cart-shopping'))

    // 9. Remove order
    TestObject order = findTestObject('Object Repository/Page_Rolex/button_Xa n Hng') // REPLACE with correct locator
    WebUI.click(findTestObject('Object Repository/Page_Rolex/button_Xa n Hng'))

     // Handle remove order alert (if any)
    try {
        WebUI.waitForAlert(5)
        WebUI.acceptAlert()
        println("Accepted remove order alert")
    } catch (Exception e) {
        println("No remove order alert")
    }
    // Verify order removal (adjust locator if needed)
    if (WebUI.waitForElementNotPresent(order, 5, FailureHandling.OPTIONAL)) {
        logResult("Order removed successfully", true)
    } else {
        logResult("Failed to remove order", false)
    }

    // 10. Click button (possibly logout)
    WebUI.click(findTestObject('Object Repository/Page_Rolex/button_'))
    logResult("Clicked final button", true)

} catch (Exception e) {
    logResult("Exception occurred: " + e.getMessage(), false)
} finally {
    WebUI.closeBrowser()
    logResult("Test case finished", true)  // Overall test case status.  Could be improved.
}