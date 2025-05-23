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
import org.openqa.selenium.interactions.Actions
import org.openqa.selenium.Keys
import com.kms.katalon.core.testobject.ConditionType
import org.openqa.selenium.JavascriptExecutor
import org.openqa.selenium.WebElement
import org.openqa.selenium.By

import stepCapture.StepCaptureHelper
StepCaptureHelper captureHelper = new StepCaptureHelper()


// Dapatkan browser yang sedang dieksekusi
def browserType = DriverFactory.getExecutedBrowser().getName()
// Cek apakah browser yang digunakan adalah Firefox
if (!browserType.toLowerCase().contains('firefox')) {
	WebUI.comment("Test case hanya dapat dijalankan di Firefox. Test dihentikan.")
	assert false : "Test dihentikan karena bukan Firefox."
}
// Jika browser adalah Firefox, lanjutkan eksekusi
WebUI.comment("Test berjalan di Firefox, melanjutkan eksekusi...")

String testCaseName = RunConfiguration.getExecutionSourceName()

// scaleSelect = auto, page-actual, page-fit, page-width, custo, 0.5, 0.75, 1, 1.25, 1.5, 2, 3, 4
def listMethodeCapture = [
	"Pemindahbukuan": [
		"stepCapture": "scroll", 
		"scaleSelect": "page-width"
	],
	"Bucket Adjusment": [
		"stepCapture": "page", 
		"scaleSelect": "page-fit"
	],
	"Asuransi": [
		"stepCapture": "page", 
		"scaleSelect": "page-fit"
	],
]

// Path ke file Excel
String excelFilePath = RunConfiguration.getProjectDir() + "/Data Files/SkenarioEnhanceRekSingleSide.xlsx"
FileInputStream file = new FileInputStream(excelFilePath)
Workbook workbook = new XSSFWorkbook(file)
Sheet sheet1 = workbook.getSheet("Batch")

WebUI.openBrowser('')
WebUI.navigateToUrl('http://192.168.174.45/Login')
WebUI.maximizeWindow()

// Loop melalui data di Sheet1
List<String> segmenAgen = ['Korporasi & Enterprise', 'Syndikasi as Fasilitas', 'Syndikasi Escrow']
for (int i = 1; i <= sheet1.getLastRowNum(); i++) {
	Row row = sheet1.getRow(i)
	String checkRunning = row.getCell(2).getStringCellValue()
	if (row != null && checkRunning != "") {
		String NoTC = row.getCell(0).getStringCellValue()
		String NoMemo = row.getCell(1).getStringCellValue()
		String IsRunning = row.getCell(2).getStringCellValue()
		String UseCase = row.getCell(3).getStringCellValue()
		String Segmen = row.getCell(4).getStringCellValue()
		String Pencairan = row.getCell(5).getStringCellValue()
		String Skenario = row.getCell(6).getStringCellValue()
		String Nominal = String.valueOf((long) row.getCell(7).getNumericCellValue())
		String RMNpp = String.valueOf((long) row.getCell(8).getNumericCellValue())
		String RMName = row.getCell(9).getStringCellValue()
		String MakerNpp = String.valueOf((long) row.getCell(10).getNumericCellValue())
		String MakerPassword = row.getCell(11).getStringCellValue()
		String MakerName = row.getCell(12).getStringCellValue()
		String MakerPositionName = row.getCell(13).getStringCellValue()
		String MakerRole = row.getCell(14).getStringCellValue()

		// Set Global Variables
		GlobalVariable.NoTC = NoTC
		GlobalVariable.NoMemo = NoMemo
		GlobalVariable.Segmen = Segmen
		GlobalVariable.Pencairan = Pencairan
		GlobalVariable.UseCase = UseCase
		GlobalVariable.Skenario = Skenario
		GlobalVariable.Nominal = Nominal
		
		String newDirectoryPath = GlobalVariable.PathCapture+"\\"+NoTC+"\\"+testCaseName
		GlobalVariable.newDirectoryPath = newDirectoryPath
		Integer numberCapture = 1
		
		File directory = new File(newDirectoryPath)
		directory.mkdirs()
		
		// Login
		WebUI.setText(findTestObject('Object Repository/Login/inputtxtUsername'), MakerNpp)
		WebUI.setText(findTestObject('Object Repository/Login/inputtxtPassword'), MakerPassword)
		WebUI.click(findTestObject('Object Repository/Login/button_Sign In'))
		WebUI.delay(3)
		
		// View Batch
		while (true) {
			try {
				// Akses halaman Monitoring Batch Success
				if (Segmen == 'BOP') {
					WebUI.waitForElementVisible(findTestObject('Object Repository/BOP/CreateNewBatch/p_Admin Kredit_Menu'), 30)
					// View Batch
					WebUI.click(findTestObject('Object Repository/BOP/CreateNewBatch/p_Admin Kredit_Menu'))
					WebUI.click(findTestObject('Object Repository/BOP/SuccessBatch/a_Admin Kredit History_BOP'))
				}
				else {
					WebUI.waitForElementVisible(findTestObject('Object Repository/COP/a_Admin Kredit'), 30)
					// View Batch
					WebUI.click(findTestObject('Object Repository/COP/a_Admin Kredit'))
					WebUI.click(findTestObject('Object Repository/COP/SuccessBatch/a_Monitoring Batch Success'))
				}
				
				// Input nomor batch
				WebUI.waitForElementVisible(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/input_filter_no_batch'), 30)
				WebUI.setText(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/input_filter_no_batch'), NoMemo)
				WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/search_button'))
				
				// Cek keberadaan elemen button_View Batch Success
				TestObject viewBatchButton = findTestObject('Object Repository/COP/SuccessBatch/button_View Batch Success')
				boolean isViewBatchExists = WebUI.verifyElementPresent(viewBatchButton, 30, FailureHandling.OPTIONAL)
				
				if (isViewBatchExists) {
					// Klik tombol View Batch Success jika ditemukan
					WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Search Batch.png')
					WebUI.click(viewBatchButton)
					WebUI.delay(3)
					break // Keluar dari loop jika proses selesai
				} else {
					if (Segmen == 'BOP') {
						WebUI.click(findTestObject('Object Repository/COP/a_Admin Kredit'))
						WebUI.click(findTestObject('Object Repository/COP/a_Monitoring Batch Progress  Failed'))
					}
					else {
						WebUI.click(findTestObject('Object Repository/BOP/CreateNewBatch/p_Admin Kredit_Menu'))
						WebUI.click(findTestObject('Object Repository/BOP/CreateNewBatch/a_Admin Kredit_SubMenu'))
					}
					// Input nomor batch ke form Monitoring Batch Progress & Failed
					WebUI.setText(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/input_filter_no_batch'), NoMemo)
					WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/search_button'))
					WebUI.delay(5)
					
					// Klik tombol View jika ditemukan
					WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Search Batch.png')
					WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_View'))
					WebUI.delay(3)
					break // Keluar dari loop jika proses selesai
				}
			} catch (Exception e) {
				// Tangani jika terjadi kesalahan dan ulangi proses dari awal
				println("Terjadi kesalahan: ${e.message}")
				continue
			}
		}
		
		CustomKeywords.'custom.CustomKeywords.captureFullPageInSections'(newDirectoryPath+'/', numberCapture++ +'. Batch Full')
		WebUI.delay(3)
		WebUI.click(findTestObject('Object Repository/COP/div_Approval History'))
		
		if (Segmen == 'BOP') {
			WebUI.scrollToElement(findTestObject('Object Repository/COP/div_Approval History'), 30)
		}
		else {
			WebUI.scrollToElement(findTestObject('Object Repository/COP/SuccessBatch/button_View Summary'), 30)
		}
		
		WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Approval History.png')
		WebUI.delay(1)
		WebUI.click(findTestObject('Object Repository/COP/SuccessBatch/button_View Result Success'))
		
		// Result PDF pada page asic
		WebDriver driver = DriverFactory.getWebDriver()
		ArrayList<String> tabs = new ArrayList<>(driver.getWindowHandles())
		driver.switchTo().window(tabs.get(1)) // Beralih ke tab baru
		WebUI.delay(3)
		WebUI.click(findTestObject('Object Repository/COP/SuccessBatch/a_Tab Result PDF'))
		WebUI.delay(3)
		
		String dynamicXPath = "//iframe[@id='pdfContent-1']"
		TestObject iframeElement = new TestObject('iframeElement')
		iframeElement.addProperty('xpath', ConditionType.EQUALS, dynamicXPath)
		WebUI.verifyElementPresent(iframeElement, 10)
		String pdfUrl = WebUI.getAttribute(iframeElement, 'src')
		println("PDF URL: " + pdfUrl)
		WebUI.delay(10)
		WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Result Page.png')
		driver.close()
		WebUI.delay(2)
		
		// Kembali ke tab sebelumnya (tab awal)
		driver.switchTo().window(tabs.get(0))
		
		// Buka tab baru menggunakan JavascriptExecutor untuk membuka result PDF
		String stepCapture = listMethodeCapture[UseCase]["stepCapture"]
		String scaleSelect = listMethodeCapture[UseCase]["scaleSelect"]
		println("stepCapture: "+stepCapture+" scaleSelect: "+scaleSelect)
		JavascriptExecutor jsExecutor = (JavascriptExecutor) driver
		jsExecutor.executeScript("window.open(arguments[0], '_blank');", "")
		WebUI.delay(2)
		ArrayList<String> tabs2 = new ArrayList<>(driver.getWindowHandles())
		driver.switchTo().window(tabs2.get(1))
		WebUI.delay(2)
		WebUI.navigateToUrl(pdfUrl)
		
		Actions actions = new Actions(driver)
		TestObject inputElement = new TestObject('dynamicInput')
		inputElement.addProperty('xpath', ConditionType.EQUALS, "//input[@id='pageNumber' and @type='number']")
		WebUI.waitForElementPresent(inputElement, 10)
		
		WebElement button = driver.findElement(By.id("sidebarToggleButton"))
		button.click()
		WebUI.delay(2)
		
		WebUI.executeJavaScript("document.querySelector('#scaleSelect').value = '"+scaleSelect+"'; document.querySelector('#scaleSelect').dispatchEvent(new Event('change'));", null)
		WebUI.delay(2)
		
		if(stepCapture == 'scroll') {
			captureHelper.scrollCapture(newDirectoryPath, numberCapture)
		}
		else if (stepCapture == 'page') {
			captureHelper.pageCapture(newDirectoryPath, numberCapture)
		}
		numberCapture++
		
		driver.close()
		WebUI.delay(5)
		driver.switchTo().window(tabs.get(0))
		
		// Logout
		WebUI.click(findTestObject('Object Repository/Login/i_User Logout'))
		WebUI.click(findTestObject('Object Repository/Login/a_Logout'))
		WebUI.delay(3)
	}
}

// Tutup
WebUI.closeBrowser()