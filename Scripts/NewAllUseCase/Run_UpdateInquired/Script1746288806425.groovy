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

import utils.LogHelper
import excel.ExcelHelper
import approval.ApprovalHelper
import logger.TestStepLogger
import custom.Select2Handler
import db.DBUtils

// Koneksi ke database
DBUtils.connectDB()

String stepName = 'Maker'

// Path ke file Excel
String excelFilePath = RunConfiguration.getProjectDir() + GlobalVariable.PathDataExcel
FileInputStream file = new FileInputStream(excelFilePath)
Workbook workbook = new XSSFWorkbook(file)
Sheet sheetBatch = workbook.getSheet("Batch")
Sheet sheetActivity = workbook.getSheet("Activity")
Sheet sheetApproval = workbook.getSheet("Approval")
Sheet sheetPemindahbukuan = workbook.getSheet("Act Pemindahbukuan")
Sheet sheetPembukaan = workbook.getSheet("Act Pembukaan Rek")

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
		
		GlobalVariable.NoTC = NoTC
		GlobalVariable.NoMemo = NoMemo
		
		// Search Approval
		def result  = ApprovalHelper.getApprovalData(
			sheetActivity,
			sheetPemindahbukuan,
			sheetPembukaan,
			sheetApproval,
			NoTC,
			Segmen
		)
		def dataApproval = result.dataApproval
		def approvalIdTerbanyak = result.maxApprovalId
		
		println "dataApproval : ${dataApproval}"
		println "Approval Id terbanyak: ${approvalIdTerbanyak}"
		
		NextApproverName_1 = dataApproval[0][2]
		println "NextApproverName_1 : ${NextApproverName_1}"
		
		// Update data Login
		String updateQuery = "update USER_LOGIN set IsLogin = '1' where Username = '${dataApproval[0][0]}'"
		println(updateQuery)
		DBUtils.executeUpdate(updateQuery)
		
//		String newDirectoryPath = GlobalVariable.PathCapture+"\\"+NoTC+"\\"+stepName
//		GlobalVariable.newDirectoryPath = newDirectoryPath
		Integer numberCapture = 9
		dirCapture = stepName
		
		// Buat TestObject dinamis untuk elemen loading
		TestObject loadingPanel = new TestObject().tap {
			addProperty("xpath", ConditionType.EQUALS, "//div[contains(@class, 'jsgrid-load-panel')]")
		}
		
		// Login
		WebUI.setText(findTestObject('Object Repository/Login/inputtxtUsername'), MakerNpp)
		WebUI.setText(findTestObject('Object Repository/Login/inputtxtPassword'), MakerPassword)
		WebUI.click(findTestObject('Object Repository/Login/button_Sign In'))
		
		WebUI.waitForElementVisible(findTestObject('Object Repository/COP/a_Admin Kredit'), 30)
		
		// View Batch
		WebUI.click(findTestObject('Object Repository/COP/a_Admin Kredit'))
		WebUI.click(findTestObject('Object Repository/COP/a_Monitoring Batch Progress  Failed'))
		WebUI.waitForElementNotVisible(loadingPanel, 30)
		WebUI.waitForElementVisible(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/search_button'), 30)
		WebUI.setText(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/input_filter_no_batch'), NoMemo)
		WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/search_button'))
		WebUI.delay(3)
		
		// Tunggu maksimal 30 detik hingga loading tidak terlihat
		WebUI.waitForElementNotVisible(loadingPanel, 30)
//		WebUI.waitForElementVisible(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_View'), 30)
		WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_View'))
		WebUI.delay(2)
		
		// Update Inquired / Inquiry Incomplete
		WebUI.scrollToElement(findTestObject('Object Repository/COP/DokUnderlying/label_Flag Batch'), 30)
		TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, numberCapture++, 1, 'Proses Inquiry selesai', dirCapture, true, false)
		
		WebDriver driver = DriverFactory.getWebDriver()
		def activityTableRows = driver.findElements(By.cssSelector('#activityTable tbody tr'))
		println("jumlah activity : "+activityTableRows.size())
		
		Boolean isExistBulkUpload = false
		List<Integer> arraySeqAct = new ArrayList<>()
		for (int j = 1; j <= sheetActivity.getLastRowNum(); j++) {
			Row rowActivity = sheetActivity.getRow(j)
			if (rowActivity != null && ExcelHelper.getCellValueAsString(rowActivity, 0) == NoTC) {
				arraySeqAct.add(ExcelHelper.getCellValueAsString(rowActivity, 2))
				if (ExcelHelper.getCellValueAsString(rowActivity, 6) == 'Y') {
					isExistBulkUpload = true
				}
			}
		}
		
		int NumberAct = 1;
		int countActivity = activityTableRows.size()
		if (isExistBulkUpload == false ) {
			countActivity = arraySeqAct.size()
		}
		
		int NumberActCapture = 1;
		for (int j = 0; j < countActivity; j++) {
			GlobalVariable.Seq = NumberAct
			if (isExistBulkUpload == false ) {
				GlobalVariable.Seq = arraySeqAct[j]
			}
			dirCapture = stepName+"/Form Inquired Activity-"+NumberAct
			
			WebElement activityTableRow = activityTableRows.get(j)
			String activityName = activityTableRow.findElements(By.tagName('td')).get(1).getText().trim()
			println("update activity ke-"+j+" : "+activityName)
			TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, NumberActCapture++, 2, "View Use Case ${activityName}", dirCapture, true, false)
		
//			int updateCount = activityName.equalsIgnoreCase('Pemindahbukuan Dana') ? 2 : 1
			int updateCount = 1
//			if(activityName.equalsIgnoreCase('Pemindahbukuan Dana') || activityName.equalsIgnoreCase('Bucket Adjustment')) {
//			if(activityName.equalsIgnoreCase('Bucket Adjustment')) {
//				updateCount = 2
//			}
//		
			for (int k = 0; k < updateCount; k++) {
				String Status = "Inquired"
				if(updateCount == 2 && k == 0) {
					 Status = "Inquiry Incomplete"
				}
				println("update status : "+Status)
				
				TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, NumberActCapture++, 3, 'Status '+Status, dirCapture, true, false)
				// Klik tombol Edit
				WebElement editButton = activityTableRow.findElement(By.cssSelector("button[onclick^='viewDetailWithSetViewed']"))
				WebUI.delay(1)
				WebUI.executeJavaScript("arguments[0].click();", [editButton])
		
				if (activityName == 'Pemindahbukuan Dana') {
					WebUI.callTestCase(findTestCase('Test Cases/NewAllUseCase/ActivityUpdate/PemindahbukuanDana'), [:])
				}
				
				if (activityName == 'Asuransi') {
					WebUI.callTestCase(findTestCase('Test Cases/NewAllUseCase/ActivityUpdate/Asuransi'), [:])
				}
				
				if (activityName == 'Bucket Adjustment') {
					WebUI.callTestCase(findTestCase('Test Cases/NewAllUseCase/ActivityUpdate/BucketAdjusment'), [:])
				}
				
				if (activityName != 'Penutupan Rek') {
					// Tunggu form dan klik tombol Update
					TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, NumberActCapture++, 3, "Update Activity ${Status}", dirCapture, true, true)
					WebUI.delay(2)
					WebUI.waitForElementVisible(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_Update'),30)
					TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, NumberActCapture++, 3, "Submit Update ${Status}", dirCapture, true, false)
					WebUI.scrollToElement(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_Update'), 30)

					WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_Update'))
					
					if (activityName == 'Bucket Adjustment') {
						TestObject AlertConfirm = findTestObject('Object Repository/ValidasiPK/AlertConfirm')
						if (WebUI.verifyElementPresent(AlertConfirm, 5, FailureHandling.OPTIONAL)) {
							TestStepLogger.addOutputWithUserAndCapture(NoTC, stepName, numberCapture++, 1, "Muncul popup konfirmasi Rekening single side menggunakan Rek Simpanan atau GL dan Dana Cukup", "Output", true, false)
							WebUI.click(AlertConfirm)
						}
					}
					
					WebUI.waitForElementVisible(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_OK'), 30)
					TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, NumberActCapture++, 3, "Berhasil Update ${Status}", dirCapture, true, false)
					WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_OK'))
				}
				else {
					WebUI.click(findTestObject('Object Repository/Activity/ActivityPenutupanRek_Object/button_OK_Alert No Such Account_Inquiry Failed'))
					WebUI.click(findTestObject('Object Repository/Activity/ActivityPenutupanRek_Object/button_Back_Form'))
				}
		
				// Tunggu kembali ke halaman utama (tabel muncul lagi)
				WebUI.scrollToElement(findTestObject('Object Repository/COP/DokUnderlying/label_Flag Batch'), 30)
				TestObject tableObject = new TestObject().addProperty("id", com.kms.katalon.core.testobject.ConditionType.EQUALS, "activityTable")
				WebUI.waitForElementVisible(tableObject, 10)
				WebUI.delay(2)
		
				// Refresh baris setelah kembali ke halaman utama
				activityTableRows = driver.findElements(By.cssSelector('#activityTable tbody tr'))
				activityTableRow = activityTableRows.get(j)
			}
			
			NumberAct++
			WebUI.delay(1)
		}
		
		dirCapture = stepName
		
		WebUI.scrollToElement(findTestObject('Object Repository/COP/DokUnderlying/label_Flag Batch'), 30)
		
		// Submit to approval
		WebUI.scrollToElement(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_Submit_Batch'), 30)
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
		
		Select2Handler.handleSelect2DropdownWithScreenshot(NoTC, stepName, numberCapture++, dirCapture, NextApproverName_1)
		
		// comment
		WebUI.setText(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/textarea__Comment'), 'submit ke approver 1')
		TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, numberCapture++, 1, "Pilih RM Pengelola, Next Approver, input comment, dan Submit Batch", dirCapture, true, false)
		WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_Submit_Batch'))
		WebUI.waitForElementVisible(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_OK_sukses_submit'), 30)
		TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, numberCapture++, 1, "Sukses Submit Batch", dirCapture, true, false)
		WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_OK_sukses_submit'))
		
		// View After Submit to approval
		WebUI.setText(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/input_filter_no_batch'), NoMemo)
		WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/search_button'))
		WebUI.waitForElementNotVisible(loadingPanel, 30)
		TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, numberCapture++, 1, "Status setelah submit batch Waiting for Approval", dirCapture, true, false)
		
		// Logout
		WebUI.click(findTestObject('Object Repository/Login/i_User Logout'))
		WebUI.click(findTestObject('Object Repository/Login/a_Logout'))
		WebUI.delay(3)
		
		// tulis log
//		LogHelper.writeLog(testCaseName, NoTC+" "+Segmen+" "+UseCase, "PASS")
	}
}

// Tutup
DBUtils.closeConnection()
workbook.close()
file.close()
WebUI.closeBrowser()