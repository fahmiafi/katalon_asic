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

import utils.LogHelper

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
		String NextApproverNpp_1 = String.valueOf((long) row.getCell(15).getNumericCellValue())
		String NextApproverPassword_1 = row.getCell(16).getStringCellValue()
		String NextApproverName_1 = row.getCell(17).getStringCellValue()
		String NextApproverPositionName_1 = row.getCell(18).getStringCellValue()
		String NextApproverRole_1 = row.getCell(19).getStringCellValue()
		
		String newDirectoryPath = GlobalVariable.PathCapture+"\\"+NoTC+"\\"+testCaseName
		GlobalVariable.newDirectoryPath = newDirectoryPath
		Integer numberCapture = 1
		
		File directory = new File(newDirectoryPath)
		directory.mkdirs()
		
		// Login
		WebUI.setText(findTestObject('Object Repository/Login/inputtxtUsername'), MakerNpp)
		WebUI.setText(findTestObject('Object Repository/Login/inputtxtPassword'), MakerPassword)
//		WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Login.png')
		WebUI.click(findTestObject('Object Repository/Login/button_Sign In'))
		
		WebUI.waitForElementVisible(findTestObject('Object Repository/COP/a_Admin Kredit'), 30)
		
		// View Batch
		WebUI.click(findTestObject('Object Repository/COP/a_Admin Kredit'))
		WebUI.click(findTestObject('Object Repository/COP/a_Monitoring Batch Progress  Failed'))
		WebUI.waitForElementVisible(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/search_button'), 30)
		WebUI.setText(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/input_filter_no_batch'), NoMemo)
		WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/search_button'))
		WebUI.delay(3)
		WebUI.waitForElementVisible(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_View'), 30)
		WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_View'))
		
		// Update Inquiry Incomplete (Pemindahbukuan)
		if (UseCase == "Pemindahbukuan") {
			WebUI.scrollToElement(findTestObject('Object Repository/COP/DokUnderlying/label_Dokumen Underlying'), 30)
			WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Inquiry Incomplete.png')
			WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_action_update'))
			CustomKeywords.'custom.CustomKeywords.captureFullPageInSections'(newDirectoryPath+'/', numberCapture++ +'. form Update Inquiry Incomplete')
			WebUI.waitForElementVisible(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_Update'),30)
			WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_Update'))
			WebUI.waitForElementVisible(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_OK'), 30)
			WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Sukses Update Inquiry Incomplete.png')
			WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_OK'))
			WebUI.delay(3)
		}
		
		// Update Inquired
		WebUI.scrollToElement(findTestObject('Object Repository/COP/DokUnderlying/label_Dokumen Underlying'), 30)
		WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Inquired.png')
		WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_action_update'))
		
		if (UseCase == 'Asuransi') {
			WebUI.selectOptionByLabel(findTestObject('Object Repository/Activity/ActivityAsuransi_Object/select_Perusahaan Asuransi'), '01 : Asuransi Tripakarta', false)
		}
		
		if (UseCase != 'Penutupan Rek') {
			CustomKeywords.'custom.CustomKeywords.captureFullPageInSections'(newDirectoryPath+'/', numberCapture++ +'. form Update Inquired')
			WebUI.waitForElementVisible(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_Update'),30)
			WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_Update'))
			WebUI.waitForElementVisible(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_OK'), 30)
			WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Sukses Update Inquired.png')
			WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_OK'))
		}
		else {
			WebUI.click(findTestObject('Object Repository/Activity/ActivityPenutupanRek_Object/button_OK_Alert No Such Account_Inquiry Failed'))
			WebUI.click(findTestObject('Object Repository/Activity/ActivityPenutupanRek_Object/button_Back_Form'))
		}
		
		
		WebUI.scrollToElement(findTestObject('Object Repository/COP/DokUnderlying/label_Dokumen Underlying'), 30)
		WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Inquired After View.png')
		
		// Submit to approval
		WebDriver driver = DriverFactory.getWebDriver()
		JavascriptExecutor jsExecutor = (JavascriptExecutor) driver
		jsExecutor.executeScript("window.scrollTo(0, document.body.scrollHeight);")
		// rm
		WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/span_Pilih RM'))
		WebUI.delay(1)
		List<WebElement> rmOptions = driver.findElements(By.xpath("//li[contains(@class,'select2-results__option') and contains(text(),'" + RMName + "')]"))
		rmOptions[0].click()
		// approver
		WebUI.click(findTestObject('Object Repository/COP/button_Refresh Approver'))
		WebUI.delay(2)
		WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/span_next approver'))
		WebUI.delay(1)
		List<WebElement> approverOptions = driver.findElements(By.xpath("//li[contains(@class,'select2-results__option')]"))
		// Pastikan ada opsi yang ditemukan
		int totalOptions = approverOptions.size()
		if (totalOptions > 0) {
			println("Jumlah total opsi dalam Select2: " + totalOptions)
		
			int displayedOptions = 6  // Jumlah opsi yang ditampilkan per halaman
			int screenshotCount = 1    // Nomor urut screenshot
			
			// **1️⃣ Ambil screenshot pertama sebelum scroll**
			WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +". Dropdown Next Approver Page_${screenshotCount}.png")
			println("Screenshot awal (tanpa scroll) berhasil disimpan.")
			
			// **2️⃣ Scroll bertahap jika opsi lebih dari 6**
			if (totalOptions > displayedOptions) {
				JavascriptExecutor js = (JavascriptExecutor) driver
				
				for (int j = displayedOptions; j < totalOptions; j += displayedOptions) {
					WebElement nextOption = approverOptions[j]
					
					// **Scroll ke opsi berikutnya**
					js.executeScript("arguments[0].scrollIntoView(true);", nextOption)
					WebUI.delay(1) // Jeda agar scroll berjalan dengan baik
					
					// **Ambil screenshot setelah scroll**
					screenshotCount++
					WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +". Dropdown Next Approver Page_${screenshotCount}.png")
					println("Screenshot halaman ke-${screenshotCount} berhasil disimpan.")
				}
			}
			
			// pilih next approver
			List<WebElement> approverOption = driver.findElements(By.xpath("//li[contains(@class,'select2-results__option') and contains(text(),'" + NextApproverName_1 + "')]"))
			approverOption[0].click()
			println("Klik opsi ${NextApproverName_1} berhasil.")
		} else {
			println("Tidak ada opsi yang ditemukan dalam Select2.")
		}
		// comment
		WebUI.setText(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/textarea__Comment'), 'submit ke approver 1')
		WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Submit To Next Approver.png')
		WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_Submit_Batch'))
		WebUI.waitForElementVisible(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_OK_sukses_submit'), 30)
		WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Sukses Submit Batch.png')
		WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_OK_sukses_submit'))
		
		// View After Submit to approval
		WebUI.setText(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/input_filter_no_batch'), NoMemo)
		WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/search_button'))
		WebUI.delay(5)
		WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Status Batch After Submit.png')
		
		// Logout
		WebUI.click(findTestObject('Object Repository/Login/i_User Logout'))
		WebUI.click(findTestObject('Object Repository/Login/a_Logout'))
		WebUI.delay(3)
		
		// tulis log
		LogHelper.writeLog(testCaseName, NoTC+" "+Segmen+" "+UseCase, "PASS")
	}
}

// Tutup
workbook.close()
file.close()
WebUI.closeBrowser()