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
				
		// Login
		WebUI.setText(findTestObject('Object Repository/Login/inputtxtUsername'), MakerNpp)
		WebUI.setText(findTestObject('Object Repository/Login/inputtxtPassword'), MakerPassword)
		WebUI.click(findTestObject('Object Repository/Login/button_Sign In'))
		
		// View Batch
		while (true) {
			try {
				WebUI.click(findTestObject('Object Repository/COP/a_Admin Kredit'))
				WebUI.click(findTestObject('Object Repository/COP/a_Monitoring Batch Progress  Failed'))
				WebUI.waitForElementVisible(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/search_button'), 30)
				WebUI.setText(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/input_filter_no_batch'), NoMemo)
				WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/search_button'))
				WebUI.delay(2)
				
				TestObject viewBatchButton = findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_View')
				boolean isViewBatchExists = WebUI.verifyElementPresent(viewBatchButton, 5, FailureHandling.OPTIONAL)
				
				if (isViewBatchExists) {
					WebUI.click(viewBatchButton)
					// Update Inquired
//					WebUI.scrollToElement(findTestObject('Object Repository/COP/DokUnderlying/button_Upload'), 30)
					
					String statusActivity = WebUI.getText(findTestObject('Object Repository/COP/td_Status Activity'))
					println("statusActivity: "+statusActivity)
					if (statusActivity == 'Execution Failed') {
						WebUI.click(findTestObject('Object Repository/COP/Button Re-Execute'))
						WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_OK_sukses_submit'))
					}
					else if (statusActivity == 'Inquiry Failed' || statusActivity == 'New') {
//						WebUI.click(findTestObject('Object Repository/COP/button_inquiry'))
						WebUI.click(findTestObject('Object Repository/COP/button_Inquiry All'))
						WebUI.click(findTestObject('Object Repository/COP/button_OK_inquiry'))
					}
					break // Keluar dari loop jika proses selesai
				}
				else {
					break
				}
			} catch (Exception e) {
				// Tangani jika terjadi kesalahan dan ulangi proses dari awal
				println("Terjadi kesalahan: ${e.message}")
				continue
			}
		}
		
		// Logout
		WebUI.click(findTestObject('Object Repository/Login/i_User Logout'))
		WebUI.click(findTestObject('Object Repository/Login/a_Logout'))
		WebUI.delay(3)
		
		// tulis log
//		LogHelper.writeLog(testCaseName, NoTC+" "+Segmen+" "+UseCase, "PASS")
	}
}

// Tutup
workbook.close()
file.close()
WebUI.closeBrowser()