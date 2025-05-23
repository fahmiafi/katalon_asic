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
String excelFilePath = RunConfiguration.getProjectDir() + "/Data Files/SkenarioEnhanceRekSingleSide.xlsx"
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
	String checkUseCase = row.getCell(3).getStringCellValue()
	if (row != null && checkRunning != "" && checkUseCase == 'Pemindahbukuan') {
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
		
		
		if (Segmen == 'BOP') {
			WebUI.waitForElementVisible(findTestObject('Object Repository/BOP/CreateNewBatch/p_Admin Kredit_Menu'), 30)
			
			// Create Batch
			WebUI.click(findTestObject('Object Repository/BOP/CreateNewBatch/p_Admin Kredit_Menu'))
			WebUI.click(findTestObject('Object Repository/BOP/CreateNewBatch/a_Admin Kredit_SubMenu'))
		}
		else {
			WebUI.waitForElementVisible(findTestObject('Object Repository/COP/a_Admin Kredit'), 30)
			
			// View Batch
			WebUI.click(findTestObject('Object Repository/COP/a_Admin Kredit'))
			WebUI.click(findTestObject('Object Repository/COP/a_Monitoring Batch Progress  Failed'))
		}
		
		WebUI.waitForElementVisible(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/search_button'), 30)
		WebUI.setText(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/input_filter_no_batch'), NoMemo)
		WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/search_button'))
		WebUI.waitForElementVisible(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_View'), 30)
		WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_View'))
		
		// Update Inquiry Incomplete (Pemindahbukuan)
//		WebUI.scrollToElement(findTestObject('Object Repository/COP/DokUnderlying/button_Upload'), 30)
		WebUI.scrollToElement(findTestObject('Object Repository/BOP/CreateNewBatch/label_List Aktivitas'), 30)
		WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Inquiry Incomplete.png')
		WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_action_update'))
		// Update dengan Bucket dan Perubahan Rek
		Sheet sheet3 = workbook.getSheet("PinbukNonPenc")
		String SelectBucket = ""
		String RekGL = ""
		String NoRekSingleSide = ""
		for (int j = 1; j <= sheet3.getLastRowNum(); j++) {
			Row rowSheetUseCase = sheet3.getRow(j)
			if (rowSheetUseCase != null && rowSheetUseCase.getCell(0).getStringCellValue() == NoTC) {
				SelectBucket = rowSheetUseCase.getCell(8).getStringCellValue()
				RekGL = rowSheetUseCase.getCell(9).getStringCellValue()
				NoRekSingleSide = rowSheetUseCase.getCell(10).getStringCellValue()
				break
			}
		}
//		println(SelectBucket)
//		println(RekGL)
//		println(NoRekSingleSide)
		
		if (Segmen != 'BOP') {
			TestObject buttonOK = findTestObject('Object Repository/Activity/ActivtyPemindahbukuan/button_OK_Jenis Pencairan Harus Diubah')
			boolean isButtonOK = WebUI.verifyElementPresent(buttonOK, 5, FailureHandling.OPTIONAL)
			if(isButtonOK) {
				WebUI.click(findTestObject('Object Repository/Activity/ActivtyPemindahbukuan/button_OK_Jenis Pencairan Harus Diubah'))
			}
		}
		
		if (SelectBucket != '') {
			WebUI.selectOptionByLabel(findTestObject('Object Repository/Activity/ActivtyPemindahbukuan/select_Pinbuk_Bucket'), SelectBucket, true)
		}
		if (RekGL != '' || NoRekSingleSide != '') {
			WebUI.click(findTestObject('Object Repository/Activity/ActivtyPemindahbukuan/input_Check_Perubahan Rekening'))
		}
		if (RekGL != '') {
			WebUI.setText(findTestObject('Object Repository/Activity/ActivtyPemindahbukuan/input_Perubahan_Debit_NoRekening'), RekGL)
		}
		if (NoRekSingleSide != '') {
			WebUI.setText(findTestObject('Object Repository/Activity/ActivtyPemindahbukuan/input_PerubahanDebit_NoRekSingleSide'), NoRekSingleSide)
		}
		
		CustomKeywords.'custom.CustomKeywords.captureFullPageInSections'(newDirectoryPath+'/', numberCapture++ +'. form Update Inquiry Incomplete')
		WebUI.waitForElementVisible(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_Update'),30)
		WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_Update'))
		WebUI.waitForElementVisible(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_OK'), 30)
		WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Sukses Update Inquiry Incomplete.png')
		WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_OK'))
		WebUI.delay(3)
		
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