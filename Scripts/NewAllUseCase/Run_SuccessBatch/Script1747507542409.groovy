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
import com.kms.katalon.core.testobject.ConditionType
import org.openqa.selenium.By
import org.openqa.selenium.WebElement
import org.openqa.selenium.JavascriptExecutor
import org.openqa.selenium.interactions.Actions

import utils.LogHelper
import excel.ExcelHelper
import stepCapture.StepCaptureHelper
import logger.TestStepLogger

// Dapatkan browser yang sedang dieksekusi
def browserType = DriverFactory.getExecutedBrowser().getName()
// Cek apakah browser yang digunakan adalah Firefox
if (!browserType.toLowerCase().contains('firefox')) {
	WebUI.comment("Test case hanya dapat dijalankan di Firefox. Test dihentikan.")
	assert false : "Test dihentikan karena bukan Firefox."
}
// Jika browser adalah Firefox, lanjutkan eksekusi
WebUI.comment("Test berjalan di Firefox, melanjutkan eksekusi...")

// scaleSelect = auto, page-actual, page-fit, page-width, custo, 0.5, 0.75, 1, 1.25, 1.5, 2, 3, 4
def listMethodeCapture = [
	"Pemindahbukuan Dana": [
		"stepCapture": "scroll",
		"scaleSelect": "page-width"
	],
	"Pembukaan Rekening Pinjaman": [
		"stepCapture": "page",
		"scaleSelect": "page-fit"
	],
	"Maintenance Rekening Pinjaman": [
		"stepCapture": "page",
		"scaleSelect": "page-fit"
	],
	"Maintenance Info Rekening Pinjaman": [
		"stepCapture": "scroll",
		"scaleSelect": "page-width"
	],
	"Restrukturisasi Rekening": [
		"stepCapture": "page",
		"scaleSelect": "page-fit"
	],
	"Bucket Adjustment": [
		"stepCapture": "page",
		"scaleSelect": "page-fit"
	],
	"Asuransi": [
		"stepCapture": "page",
		"scaleSelect": "page-fit"
	],
	"Blokir Rekening": [
		"stepCapture": "scroll",
		"scaleSelect": "page-width"
	],
]

StepCaptureHelper captureHelper = new StepCaptureHelper()
String stepName = 'Result'
String dirCapture = stepName

// Path ke file Excel
String excelFilePath = RunConfiguration.getProjectDir() + GlobalVariable.PathDataExcel
FileInputStream file = new FileInputStream(excelFilePath)
Workbook workbook = new XSSFWorkbook(file)
Sheet sheetBatch = workbook.getSheet("Batch")
Sheet sheetActivity = workbook.getSheet("Activity")

WebUI.openBrowser('')
WebUI.navigateToUrl('http://192.168.174.45/Login')
WebUI.maximizeWindow()

// Loop melalui data di Sheet1
for (int i = 1; i <= sheetBatch.getLastRowNum(); i++) {
	Row row = sheetBatch.getRow(i)
	String checkRunning = row.getCell(2).getStringCellValue()
	if (row != null && checkRunning == "Y") {
		String NoTC = ExcelHelper.getCellValueAsString(row, 0)
		String NoMemo = ExcelHelper.getCellValueAsString(row, 1)
		String IsRunning = ExcelHelper.getCellValueAsString(row, 2)
		String Segmen = ExcelHelper.getCellValueAsString(row, 3)
		String SkenarioBatch = ExcelHelper.getCellValueAsString(row, 4)
		String RMNpp = ExcelHelper.getCellValueAsString(row, 5)
		String RMName = ExcelHelper.getCellValueAsString(row, 6)
		String MakerNpp = ExcelHelper.getCellValueAsString(row, 7)
		String MakerPassword = ExcelHelper.getCellValueAsString(row, 8)
		String MakerName = ExcelHelper.getCellValueAsString(row, 9)
		String MakerPositionName = ExcelHelper.getCellValueAsString(row, 10)
		String MakerRole = ExcelHelper.getCellValueAsString(row, 11)		
		
		// Set Global Variables
		GlobalVariable.NoTC = NoTC
		GlobalVariable.NoMemo = NoMemo
		
		String newDirectoryPath = GlobalVariable.PathCapture+"\\"+NoTC+"\\"+stepName
		GlobalVariable.newDirectoryPath = newDirectoryPath
		Integer numberCapture = 1
		
		File directory = new File(newDirectoryPath)
		directory.mkdirs()
		
		// Login
		WebUI.setText(findTestObject('Object Repository/Login/inputtxtUsername'), MakerNpp)
		WebUI.setText(findTestObject('Object Repository/Login/inputtxtPassword'), MakerPassword)
//		WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Login sebagai maker.png')
		WebUI.click(findTestObject('Object Repository/Login/button_Sign In'))
		
		WebUI.waitForElementVisible(findTestObject('Object Repository/COP/a_Admin Kredit'), 30)
		
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
//					WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Status Batch Success.png')
					TestStepLogger.addOutputWithUserAndCapture(NoTC, stepName, numberCapture++, 1, 'Status Batch Success', dirCapture, true, false)
					WebUI.click(viewBatchButton)
					WebUI.delay(3)
					break // Keluar dari loop jika proses selesai
				} else {
					if (Segmen == 'BOP') {
						WebUI.click(findTestObject('Object Repository/COP/a_Admin Kredit'))
						WebUI.click(findTestObject('Object Repository/COP/a_Monitoring Batch Progress  Failed'))
					}
					else {
						WebUI.click(findTestObject('Object Repository/COP/a_Admin Kredit'))
						WebUI.click(findTestObject('Object Repository/COP/a_Monitoring Batch Progress  Failed'))
					}
					// Input nomor batch ke form Monitoring Batch Progress & Failed
					WebUI.setText(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/input_filter_no_batch'), NoMemo)
					WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/search_button'))
					WebUI.delay(5)
					
					// Klik tombol View jika ditemukan
					TestStepLogger.addOutputWithUserAndCapture(NoTC, stepName, numberCapture++, 1, 'Status Batch Success', dirCapture, true, false)
//					WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Status Batch Success.png')
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
		
		TestStepLogger.addOutputWithUserAndCapture(NoTC, stepName, numberCapture++, 1, 'Status Activity Success', dirCapture, true, true)
//		CustomKeywords.'custom.CustomKeywords.captureFullPageInSections'(newDirectoryPath+'/', numberCapture++ +'. Status Activity Success')
		WebUI.delay(3)
		WebUI.click(findTestObject('Object Repository/COP/div_Approval History'))
		
		if (Segmen == 'BOP') {
			WebUI.scrollToElement(findTestObject('Object Repository/COP/div_Approval History'), 30)
		}
		else {
			WebUI.scrollToElement(findTestObject('Object Repository/COP/div_Approval History'), 30)
		}
		
//		TestStepLogger.addOutputWithUserAndCapture(NoTC, stepName, numberCapture++, 1, 'Status Activity Success', dirCapture, true, true)
//		WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Approval History.png')
		WebUI.delay(1)
//		WebUI.click(findTestObject('Object Repository/COP/SuccessBatch/button_View Result Success'))
		
		
		WebDriver driver = DriverFactory.getWebDriver()
		def activityTableRows = driver.findElements(By.cssSelector('#activityTable tbody tr'))
		println("jumlah activity : "+activityTableRows.size())
		
		
		int NumberAct = 1;
		for (int j = 0; j < activityTableRows.size(); j++) {
			newDirectoryPath = GlobalVariable.PathCapture+"\\"+NoTC+"\\"+stepName+"\\Result Activity-"+NumberAct
			GlobalVariable.newDirectoryPath = newDirectoryPath
			
			WebElement activityTableRow = activityTableRows.get(j)
			String activityName = activityTableRow.findElements(By.tagName('td')).get(1).getText().trim()
			println("Result activity ke-"+j+" : "+activityName)
		
			// Klik tombol Result
			WebElement resultButton = activityTableRow.findElement(By.cssSelector("button[onclick^='viewResult']"))
			WebUI.delay(1)
			WebUI.executeJavaScript("arguments[0].click();", [resultButton])
	
			// ===========================================================
			// Result PDF pada page asic
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
			WebUI.delay(15)
//			WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Result Rekonsiliasi.png')
			TestStepLogger.addOutputWithUserAndCapture(NoTC, stepName, numberCapture++, 4, 'Result Rekonsiliasi', dirCapture, true, false)
			driver.close()
			WebUI.delay(2)
			
			// Kembali ke tab sebelumnya (tab awal)
			driver.switchTo().window(tabs.get(0))
			
			// Buka tab baru menggunakan JavascriptExecutor untuk membuka result PDF
			String stepCapture = listMethodeCapture[activityName]["stepCapture"]
			String scaleSelect = listMethodeCapture[activityName]["scaleSelect"]
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
				captureHelper.scrollCapture(stepName, NoTC, numberCapture, 'Evidence Result Rekonsiliasi')
			}
			else if (stepCapture == 'page') {
				captureHelper.pageCapture(stepName, NoTC, numberCapture, 'Evidence Result Rekonsiliasi')
			}
			numberCapture++
			
			driver.close()
			WebUI.delay(5)
			driver.switchTo().window(tabs.get(0))
			// ===========================================================
	
			// Tunggu kembali ke halaman utama (tabel muncul lagi)
			WebUI.scrollToElement(findTestObject('Object Repository/COP/DokUnderlying/label_Flag Batch'), 30)
			TestObject tableObject = new TestObject().addProperty("id", com.kms.katalon.core.testobject.ConditionType.EQUALS, "activityTable")
			WebUI.waitForElementVisible(tableObject, 10)
			WebUI.delay(2)
	
			// Refresh baris setelah kembali ke halaman utama
			activityTableRows = driver.findElements(By.cssSelector('#activityTable tbody tr'))
			activityTableRow = activityTableRows.get(j)
			
			NumberAct++
			WebUI.delay(1)
		}
		
		// Logout
		WebUI.click(findTestObject('Object Repository/Login/i_User Logout'))
		WebUI.click(findTestObject('Object Repository/Login/a_Logout'))
		WebUI.delay(2)
		
		// tulis log
//		LogHelper.writeLog(testCaseName, NoTC+" "+Segmen+" "+UseCase, "PASS")
	}
}

// Tutup
workbook.close()
file.close()
WebUI.closeBrowser()