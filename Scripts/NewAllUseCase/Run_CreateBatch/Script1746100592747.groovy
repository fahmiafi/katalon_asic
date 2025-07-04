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
import logger.TestStepLogger

String stepName = 'Maker'
GlobalVariable.stepName = stepName
String pathBulkUploadExcel = 'D:\\BNI\\RPA\\CR Antasena, Enhance Robot\\Data Excel\\'
String dirCapture = ""
dirCapture = stepName

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
				
		Integer numberCapture = 1
		dirCapture = stepName
		
		// Login
		WebUI.setText(findTestObject('Object Repository/Login/inputtxtUsername'), MakerNpp)
		WebUI.setText(findTestObject('Object Repository/Login/inputtxtPassword'), MakerPassword)
		TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, numberCapture++, 'Login sebagai maker', dirCapture, true, false)
		WebUI.click(findTestObject('Object Repository/Login/button_Sign In'))
		
		WebUI.waitForElementVisible(findTestObject('Object Repository/COP/a_Admin Kredit'), 30)
		
		// Create Batch
		WebUI.click(findTestObject('Object Repository/COP/a_Admin Kredit'))
		TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, numberCapture, 'Akses menu Admin Kredit >> Monitoring Batch Progress & Failed', dirCapture, false, false)
		WebUI.click(findTestObject('Object Repository/COP/a_Monitoring Batch Progress  Failed'))
		// Buat TestObject dinamis untuk elemen loading
		TestObject loadingPanel = new TestObject().tap {
			addProperty("xpath", ConditionType.EQUALS, "//div[contains(@class, 'jsgrid-load-panel')]")
		}
		
		TestObject loadingActivityTable = new TestObject().tap {
			addProperty("xpath", ConditionType.EQUALS, "//div[contains(@id, 'activityTable_processing')]")
		}
		
		// Tunggu maksimal 30 detik hingga loading tidak terlihat
		WebUI.waitForElementNotVisible(loadingPanel, 30)
		WebUI.waitForElementVisible(findTestObject('Object Repository/COP/a_Create New Batch'), 30)
		TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, numberCapture++, 'Create New Batch', dirCapture, true, false)
		WebUI.click(findTestObject('Object Repository/COP/a_Create New Batch'))
		
		String NamaDebitur = NoTC+' '+ SkenarioBatch
		// Potong jika panjang lebih dari 100 karakter
		if (NamaDebitur.length() > 99) {
			NamaDebitur = NamaDebitur.substring(0, 99)
		}
		WebUI.waitForElementVisible(findTestObject('Object Repository/COP/DokUnderlying/select_Kategori Underlying'), 30)
		WebUI.setText(findTestObject('Object Repository/COP/input_NoTestKey'), NoMemo)
		WebUI.setText(findTestObject('Object Repository/COP/input_NamaDebitur'), NamaDebitur.replaceAll(/\s+$/, '').replaceAll(/\s+/, ' '))
		WebUI.verifyOptionPresentByValue(findTestObject('Object Repository/COP/DokUnderlying/select_Kategori Underlying'), '1', false, 30)
		WebUI.selectOptionByValue(findTestObject('Object Repository/COP/DokUnderlying/select_Kategori Underlying'), '1', true)
		WebUI.setText(findTestObject('Object Repository/COP/DokUnderlying/inputnama'), '123')
		
		TestObject uploadField = findTestObject('Object Repository/COP/DokUnderlying/input_Dokumen Underlying_telexfile')
		String filePath = 'D:\\BNI\\RPA\\.TESTING.pdf'
		WebUI.uploadFile(uploadField, filePath)
		TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, numberCapture++, 'Input No. Memo, Nama Debitur dan Upload Dokumen Underlying', dirCapture, true, false)
		WebUI.click(findTestObject('Object Repository/COP/DokUnderlying/button_Upload'))
//		WebUI.delay(7)
		WebUI.waitForElementVisible(findTestObject('Object Repository/COP/DokUnderlying/td_No_Memo_Underlying'), 30)
		
		// Add Activity
		WebUI.scrollToElement(findTestObject('Object Repository/COP/DokUnderlying/label_Flag Batch'), 30)
		TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, numberCapture++, 'Add Activity', dirCapture, true, false)
		WebUI.click(findTestObject('Object Repository/COP/button_Add Activity'))
		int NumberAct = 1;
		for (int j = 1; j <= sheetActivity.getLastRowNum(); j++) {
			Row rowActivity = sheetActivity.getRow(j)
			String checkTcAct = ExcelHelper.getCellValueAsString(rowActivity, 0)
			if (rowActivity != null && checkTcAct == NoTC) {
				if(NumberAct != 1) {
					WebUI.scrollToElement(findTestObject('Object Repository/COP/DokUnderlying/label_Flag Batch'), 30)
					WebUI.click(findTestObject('Object Repository/COP/button_Add Activity'))
				}
				String Seq = ExcelHelper.getCellValueAsString(rowActivity, 2)
				String UseCase = ExcelHelper.getCellValueAsString(rowActivity, 3)
				String SkenarioActivity = ExcelHelper.getCellValueAsString(rowActivity, 4)
				String Pencairan = ExcelHelper.getCellValueAsString(rowActivity, 5)
				String BulkUpload = ExcelHelper.getCellValueAsString(rowActivity, 6)
				String ExcelFilename = ExcelHelper.getCellValueAsString(rowActivity, 7)
				GlobalVariable.Seq = Seq
				GlobalVariable.Pencairan = Pencairan
				GlobalVariable.UseCase = UseCase
				GlobalVariable.BulkUpload = BulkUpload
				GlobalVariable.ExcelFilename = pathBulkUploadExcel+ExcelFilename
				
				println("Menjalankan Skenario : " +NoTC+" - "+Seq+" - "+UseCase+" - "+SkenarioActivity+" - "+Pencairan)
				WebUI.delay(2)
				
				dirCapture = stepName+"/Form Activity-"+NumberAct
				GlobalVariable.newDirectoryPath = dirCapture
				
				// Panggil test case Use Case
				TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, numberCapture, 'Use Case '+UseCase, dirCapture, false, false)
				if (UseCase == "Pemindahbukuan") {
					WebUI.callTestCase(findTestCase('Test Cases/NewAllUseCase/Activity/PemindahbukuanDana'), [:])
				} else if (UseCase == "Pembukaan Rek") {
					WebUI.callTestCase(findTestCase('Test Cases/NewAllUseCase/Activity/PembukaanRek'), [:])
				} else if (UseCase == "Blokir Rekening") {
					WebUI.callTestCase(findTestCase('Test Cases/NewAllUseCase/Activity/BlokirRek'), [:])
				} else if (UseCase == "Pending Rekening") {
					WebUI.callTestCase(findTestCase('Test Cases/NewAllUseCase/Activity/PendingRek'), [:])
				} else if (UseCase == "Restrukturisasi Rek") {
					WebUI.callTestCase(findTestCase('Test Cases/NewAllUseCase/Activity/RestrukturisasiRek'), [:])
				} else if (UseCase == "Maintenance Info") {
					WebUI.callTestCase(findTestCase('Test Cases/NewAllUseCase/Activity/MaintenanceInfoRek'), [:])
				} else if (UseCase == "Maintenance Rek") {
					WebUI.callTestCase(findTestCase('Test Cases/NewAllUseCase/Activity/MaintenanceRek'), [:])
				} else if (UseCase == "Asuransi") {
					WebUI.callTestCase(findTestCase('Test Cases/NewAllUseCase/Activity/Asuransi'), [:])
				} else if (UseCase == "Penutupan Rek") {
					WebUI.callTestCase(findTestCase('Test Cases/NewAllUseCase/Activity/PenutupanRek'), [:])
				} else if (UseCase == "Bucket Adjusment") {
					WebUI.callTestCase(findTestCase('Test Cases/NewAllUseCase/Activity/BucketAdjusment'), [:])
				}
				
				WebUI.waitForElementNotVisible(loadingActivityTable, 30)
				NumberAct++
			}
		}
		
		dirCapture = stepName
		GlobalVariable.newDirectoryPath = dirCapture
		
		WebUI.scrollToElement(findTestObject('Object Repository/COP/DokUnderlying/label_Flag Batch'), 30)
		// inquiry
		WebUI.waitForElementVisible(findTestObject('Object Repository/COP/button_inquiry'), 30)
		TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, numberCapture++, 'Activity berhasil ditambahkan, status New', dirCapture, true, false)
		TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, numberCapture++, 'Inquiry pada table List Aktivitas kolom Action atau Inquiry All', dirCapture, true, false)
		WebUI.click(findTestObject('Object Repository/COP/button_Inquiry All'))
		WebUI.waitForElementVisible(findTestObject('Object Repository/COP/button_OK_inquiry'), 30)
		TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, numberCapture++, 'Inquiry berhasil', dirCapture, true, false)
		WebUI.click(findTestObject('Object Repository/COP/button_OK_inquiry'))
		WebUI.delay(1)
		TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, numberCapture++, 'Status Inquiring', dirCapture, true, false)
		
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