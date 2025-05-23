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

import org.openqa.selenium.WebDriver
import com.kms.katalon.core.webui.driver.DriverFactory
import com.kms.katalon.core.configuration.RunConfiguration
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileInputStream
import org.openqa.selenium.JavascriptExecutor
import org.openqa.selenium.By
import org.openqa.selenium.WebElement
import com.kms.katalon.core.testobject.ConditionType


WebUI.openBrowser('')
WebUI.navigateToUrl('http://192.168.174.45/Login')
WebUI.maximizeWindow()

WebDriver driver = DriverFactory.getWebDriver()
// Login
WebUI.setText(findTestObject('Object Repository/Login/inputtxtUsername'), '13538')
WebUI.setText(findTestObject('Object Repository/Login/inputtxtPassword'), 'bnitest123')
WebUI.click(findTestObject('Object Repository/Login/button_Sign In'))

WebUI.waitForElementVisible(findTestObject('Object Repository/COP/a_Admin Kredit'), 30)

// Search Batch
WebUI.click(findTestObject('Object Repository/COP/a_Admin Kredit'))
WebUI.click(findTestObject('Object Repository/COP/Approval_Object/a_Approval Cop'))
WebUI.delay(3)
// Hitung jumlah tombol "Action" yang ada di dalam tabel
String xpathButtonAction = "//table[contains(@class, 'jsgrid-table')]//button[contains(., 'Action')]"
List<WebElement> actionButtons = WebUI.findWebElements(new TestObject().addProperty('xpath', ConditionType.EQUALS, xpathButtonAction), 10)

// Cetak jumlah tombol action yang ditemukan
int jumlahButton = actionButtons.size()
println("Jumlah tombol Action yang ditemukan: " + jumlahButton)

String tableXPath = "//table[contains(@class, 'jsgrid-table')]"
for (int i = 0; i < jumlahButton; i++) {
	String buttonXPath = tableXPath + "/tbody/tr[1]/td[last()]/div/button"
	WebElement actionButton = driver.findElement(By.xpath(buttonXPath))
	actionButton.click()
	
	WebUI.setText(findTestObject('Object Repository/COP/Approval_Object/textarea_Approval_Comment'), "reject")
	
	// Submit Batch
	WebUI.click(findTestObject('Object Repository/COP/Approval_Object/button_Submit_Batch_Approval'))
	WebUI.waitForElementVisible(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_OK_sukses_submit'), 30)
	WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_OK_sukses_submit'))
	WebUI.delay(2)
}

// Tutup browser setelah selesai
WebUI.closeBrowser()