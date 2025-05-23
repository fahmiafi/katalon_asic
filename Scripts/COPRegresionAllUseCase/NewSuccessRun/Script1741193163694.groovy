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


String testCaseName = RunConfiguration.getExecutionSourceName()

// Path ke file Excel
String excelFilePath = RunConfiguration.getProjectDir() + "/Data Files/NewSkenarioAllActivity.xlsx"
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
	if (row != null && checkRunning != "" && i == 38) {
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
				WebUI.click(findTestObject('Object Repository/COP/a_Admin Kredit'))
				WebUI.click(findTestObject('Object Repository/COP/a_Monitoring Batch Success'))
				WebUI.delay(3)
				
				// Input nomor batch
				WebUI.setText(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/input_filter_no_batch'), NoMemo)
				WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/search_button'))
				WebUI.delay(5)
				
				// Cek keberadaan elemen button_View Batch Success
				TestObject viewBatchButton = findTestObject('Object Repository/COP/Page_BNI.RPA.CORE/button_View Batch Success')
				boolean isViewBatchExists = WebUI.verifyElementPresent(viewBatchButton, 5, FailureHandling.OPTIONAL)
				
				if (isViewBatchExists) {
					// Klik tombol View Batch Success jika ditemukan
					WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Search Batch.png')
					WebUI.click(viewBatchButton)
					WebUI.delay(3)
					break // Keluar dari loop jika proses selesai
				} else {
					// Jika elemen tidak ditemukan, jalankan proses Monitoring Batch Progress & Failed
					WebUI.click(findTestObject('Object Repository/COP/a_Admin Kredit'))
					WebUI.click(findTestObject('Object Repository/COP/a_Monitoring Batch Progress  Failed'))
					WebUI.delay(3)
					
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
		WebUI.scrollToElement(findTestObject('Object Repository/COP/Page_BNI.RPA.CORE/button_View Summary'), 30)
		WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Approval History.png')
		WebUI.delay(10)
		WebUI.click(findTestObject('Object Repository/COP/Page_BNI.RPA.CORE/button_View Result Success'))
		
		// Result PDF pada page asic
		WebDriver driver = DriverFactory.getWebDriver()
		ArrayList<String> tabs = new ArrayList<>(driver.getWindowHandles())
		driver.switchTo().window(tabs.get(1)) // Beralih ke tab baru
		WebUI.delay(3)
		WebUI.click(findTestObject('Object Repository/COP/Page_BNI.RPA.CORE/a_Tab Result PDF'))
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
		
		int[] PageCaptureResult
		if (UseCase == "Pemindahbukuan") {
			WebUI.delay(2)
			WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. PDF Result page_1.png')
			actions.sendKeys(Keys.PAGE_DOWN).perform()
			WebUI.delay(2)
			WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. PDF Result page_2.png')
		} else if (UseCase == "Pembukaan Rek") {
			PageCaptureResult = [1, 2, 3, 4, 8, 9, 10, 11, 12]
			for (int j = 0; j < PageCaptureResult.length; j++) {
				int page = PageCaptureResult[j]
				WebUI.click(inputElement)
				actions.keyDown(Keys.CONTROL).sendKeys("a").keyUp(Keys.CONTROL).sendKeys(Keys.DELETE).perform()
				WebUI.delay(1)
				WebUI.setText(inputElement, page.toString())
				actions.sendKeys(Keys.ENTER).perform()
				WebUI.delay(2)
				WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. PDF Result page_'+ page +'.png')
			}
		}
		
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