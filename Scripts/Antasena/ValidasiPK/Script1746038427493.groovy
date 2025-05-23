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

import utils.LogHelper
import excel.ExcelHelper

String testCaseName = RunConfiguration.getExecutionSourceName()

// Path ke file Excel
String excelFilePath = RunConfiguration.getProjectDir() + "/Data Files/SkenarioValidasiPK.xlsx"
FileInputStream file = new FileInputStream(excelFilePath)
Workbook workbook = new XSSFWorkbook(file)
Sheet sheetSkenario = workbook.getSheet("Skenario")

WebUI.openBrowser('')
WebUI.navigateToUrl('http://192.168.174.45/Login')
WebUI.maximizeWindow()

// Loop melalui data di Sheet1
for (int i = 1; i <= sheetSkenario.getLastRowNum(); i++) {
	Row row = sheetSkenario.getRow(i)
	String checkRunning = row.getCell(1).getStringCellValue()
	if (row != null && checkRunning == "Y") {
		String NoTC = ExcelHelper.getCellValueAsString(row, 0)
		String IsRunning = ExcelHelper.getCellValueAsString(row, 1)
		String UseCase = ExcelHelper.getCellValueAsString(row, 2)
		String Metode = ExcelHelper.getCellValueAsString(row, 3)
		String Skenario = ExcelHelper.getCellValueAsString(row, 4)
		String TestStep = ExcelHelper.getCellValueAsString(row, 5)
		String ExpectedResult = ExcelHelper.getCellValueAsString(row, 6)
		String DataTest = ExcelHelper.getCellValueAsString(row, 7)
		String NomorPKAwal = ExcelHelper.getCellValueAsString(row, 8)
		String NomorPKAkhir = ExcelHelper.getCellValueAsString(row, 9)
		String TanggalPKAwal = ExcelHelper.getCellValueAsString(row, 10)
		String TanggalPKAkhir = ExcelHelper.getCellValueAsString(row, 11)
		String MakerNpp = ExcelHelper.getCellValueAsString(row, 12)
		String MakerPassword = ExcelHelper.getCellValueAsString(row, 13)
		
		String newDirectoryPath = GlobalVariable.PathCapture+"\\"+NoTC+"\\"+testCaseName
		GlobalVariable.newDirectoryPath = newDirectoryPath
		Integer numberCapture = 1
		
		// Login
		WebUI.setText(findTestObject('Object Repository/Login/inputtxtUsername'), MakerNpp)
		WebUI.setText(findTestObject('Object Repository/Login/inputtxtPassword'), MakerPassword)
		WebUI.click(findTestObject('Object Repository/Login/button_Sign In'))
		
		WebUI.waitForElementVisible(findTestObject('Object Repository/COP/a_Admin Kredit'), 30)
		
		// View Batch
		WebUI.click(findTestObject('Object Repository/COP/a_Admin Kredit'))
		WebUI.click(findTestObject('Object Repository/COP/a_Monitoring Batch Progress  Failed'))
		WebUI.waitForElementVisible(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/search_button'), 30)
		WebUI.setText(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/input_filter_no_batch'), DataTest)
		WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/search_button'))
		WebUI.delay(3)
		// Buat TestObject dinamis untuk elemen loading
		TestObject loadingPanel = new TestObject().tap {
			addProperty("xpath", ConditionType.EQUALS, "//div[contains(@class, 'jsgrid-load-panel')]")
		}
		
		// Tunggu maksimal 30 detik hingga loading tidak terlihat
		WebUI.waitForElementNotVisible(loadingPanel, 30)
		WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_View'))
		WebUI.delay(2)
		
		
		// Update Inquired / Inquiry Incomplete
		WebUI.scrollToElement(findTestObject('Object Repository/COP/DokUnderlying/label_Flag Batch'), 30)
		WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. PROCEDURE - Activity status Inquired.png')
		WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_action_update'))
		
		if (UseCase == 'Maintenance Rek') {
			WebUI.scrollToElement(findTestObject('Object Repository/Activity/ActivityMaintenanceRek_Object/check_JatuhTempo'), 30)
			println("TanggalPKAwal :"+TanggalPKAwal)
			WebUI.setText(findTestObject('Object Repository/Activity/ActivityMaintenanceRek_Object/input_Tanggal PK Awal Menjadi'), TanggalPKAwal != null ? TanggalPKAwal : '')
			println("TanggalPKAkhir :"+TanggalPKAkhir)
			WebUI.setText(findTestObject('Object Repository/Activity/ActivityMaintenanceRek_Object/input_Tanggal PK Akhir Menjadi'), TanggalPKAkhir != null ? TanggalPKAkhir : '')
			println("NomorPKAwal :"+NomorPKAwal)
			WebUI.setText(findTestObject('Object Repository/Activity/ActivityMaintenanceRek_Object/input_Nomor PK Awal Menjadi'), NomorPKAwal != null ? NomorPKAwal : '')
			println("NomorPKAkhir :"+NomorPKAkhir)
			WebUI.setText(findTestObject('Object Repository/Activity/ActivityMaintenanceRek_Object/input_Nomor PK Akhir Menjadi'), NomorPKAkhir != null ? NomorPKAkhir : '')
			
			CustomKeywords.'custom.CustomKeywords.scrollToTop'()
			WebUI.click(findTestObject('Object Repository/Activity/ActivityMaintenanceRek_Object/input_Nomor Rekening'))
		}
		else if (UseCase == 'Maintenance Info') {
			println("NomorPKAwal :"+NomorPKAwal)
			WebUI.setText(findTestObject('Object Repository/Activity/ActivityMaintInfoRek_Object/input_Nomor PK Awal Menjadi'), NomorPKAwal != null ? NomorPKAwal : '')
			println("NomorPKAkhir :"+NomorPKAkhir)
			WebUI.setText(findTestObject('Object Repository/Activity/ActivityMaintInfoRek_Object/input_Nomor PK Akhir Menjadi'), NomorPKAkhir != null ? NomorPKAkhir : '')
			println("TanggalPKAwal :"+TanggalPKAwal)
			WebUI.setText(findTestObject('Object Repository/Activity/ActivityMaintInfoRek_Object/input_Tanggal PK Awal Menjadi'), TanggalPKAwal != null ? TanggalPKAwal : '')
			println("TanggalPKAkhir :"+TanggalPKAkhir)
			WebUI.setText(findTestObject('Object Repository/Activity/ActivityMaintInfoRek_Object/input_Tanggal PK Akhir Menjadi'), TanggalPKAkhir != null ? TanggalPKAkhir : '')
			
			CustomKeywords.'custom.CustomKeywords.scrollToTop'()
			WebUI.click(findTestObject('Object Repository/Activity/ActivityMaintInfoRek_Object/input_Nomor Rekening'))
		}
		CustomKeywords.'custom.CustomKeywords.captureFullPageInSections'(newDirectoryPath+'/', numberCapture++ +'. PROCEDURE - Update Nomor PK dan Tanggal PK')
		
		WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. PROCEDURE - Submit Update.png')
		WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_Update'))
		
		TestObject AlertTitle = findTestObject('Object Repository/ValidasiPK/AlertTitle')
		TestObject AlertMessage = findTestObject('Object Repository/ValidasiPK/AlertMessage')
		TestObject AlertConfirm = findTestObject('Object Repository/ValidasiPK/AlertConfirm')
		// Tunggu sampai alert muncul
		WebUI.waitForElementVisible(AlertTitle, 10)
		
		// Ambil teks pesan alert
		String Title = WebUI.getText(AlertTitle)
		String Message = WebUI.getText(AlertMessage)
		println("Title : " + Title)
		println("Message :" + Message)
		WebUI.delay(3)
		WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. PROCEDURE - Muncul popup message.png')
		if (Title == 'Success') {
			WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. OUTPUT - Muncul popup message Success.png')
		}
		else {
			WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. OUTPUT - Muncul popup message Error.png')
		}
		WebUI.click(AlertConfirm)
		
		// Logout
		WebUI.click(findTestObject('Object Repository/Login/i_User Logout'))
		WebUI.click(findTestObject('Object Repository/Login/a_Logout'))
		WebUI.delay(2)
	}
}

// Tutup
workbook.close()
file.close()
WebUI.closeBrowser()