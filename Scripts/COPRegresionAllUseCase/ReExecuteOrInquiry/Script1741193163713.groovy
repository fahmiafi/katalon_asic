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

List<String> arrTc = ['TC-031', 'TC-037', 'TC-051', 'TC-091', 'TC-092', 'TC-093', 'TC-094', 'TC-095']
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
		
		// Login
		WebUI.setText(findTestObject('Object Repository/Login/inputtxtUsername'), MakerNpp)
		WebUI.setText(findTestObject('Object Repository/Login/inputtxtPassword'), MakerPassword)
//		WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Login.png')
		WebUI.click(findTestObject('Object Repository/Login/button_Sign In'))
		WebUI.delay(3)
		
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
					else if (statusActivity == 'Inquiry Failed') {
						WebUI.click(findTestObject('Object Repository/COP/button_inquiry'))
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
	}
}

// Tutup
workbook.close()
file.close()
WebUI.closeBrowser()