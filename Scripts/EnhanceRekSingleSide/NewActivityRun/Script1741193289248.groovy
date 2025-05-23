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
		WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Login.png')
		WebUI.click(findTestObject('Object Repository/Login/button_Sign In'))
		
		if (Segmen == 'BOP') {
			println("Menjalankan Skenario BOP")
			WebUI.callTestCase(findTestCase('Test Cases/EnhanceRekSingleSide/CreateNewBatchBOP'), [:])
		}
		else {
			println("Menjalankan Skenario COP")
			WebUI.callTestCase(findTestCase('Test Cases/EnhanceRekSingleSide/CreateNewBatchCOP'), [:])
		}
		
		
		
		// Panggil test case Use Case
		if (UseCase == "Pemindahbukuan") {
			WebUI.callTestCase(findTestCase('Test Cases/EnhanceRekSingleSide/Activity/PemindahbukuanDana'), [:])
		} else if (UseCase == "Asuransi") {
			WebUI.callTestCase(findTestCase('Test Cases/EnhanceRekSingleSide/Activity/Asuransi'), [:])
		} else if (UseCase == "Bucket Adjusment") {
			WebUI.callTestCase(findTestCase('Test Cases/EnhanceRekSingleSide/Activity/BucketAdjusment'), [:])
		}
		
		if (Segmen == 'BOP') {
			WebUI.scrollToElement(findTestObject('Object Repository/BOP/CreateNewBatch/label_List Aktivitas'), 30)
		}
		else {
			WebUI.scrollToElement(findTestObject('Object Repository/COP/DokUnderlying/button_Upload'), 30)
		}
		
		// inquiry
		WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Activity New.png')
		WebUI.click(findTestObject('Object Repository/COP/button_inquiry'))
		WebUI.waitForElementVisible(findTestObject('Object Repository/COP/button_OK_inquiry'), 30)
		WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Alert Inquiry.png')
		WebUI.click(findTestObject('Object Repository/COP/button_OK_inquiry'))
		WebUI.delay(1)
		WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Inquiring.png')
		
		// Logout
		WebUI.click(findTestObject('Object Repository/Login/i_User Logout'))
		WebUI.click(findTestObject('Object Repository/Login/a_Logout'))
		WebUI.delay(2)
		
		// tulis log
		LogHelper.writeLog(testCaseName, NoTC+" "+Segmen+" "+UseCase, "PASS")
	}
}

// Tutup
workbook.close()
file.close()
WebUI.closeBrowser()