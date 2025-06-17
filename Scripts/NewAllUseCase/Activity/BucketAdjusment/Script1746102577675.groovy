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

import com.kms.katalon.core.configuration.RunConfiguration
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileInputStream
import org.openqa.selenium.WebDriver
import com.kms.katalon.core.webui.driver.DriverFactory
import org.openqa.selenium.By
import org.openqa.selenium.WebElement
import excel.ExcelHelper
import custom.ActivityUtils
import logger.TestStepLogger

String NoTC = GlobalVariable.NoTC
String Seq = GlobalVariable.Seq
String Pencairan = GlobalVariable.Pencairan
String UseCase = GlobalVariable.UseCase
String newDirectoryPath = GlobalVariable.newDirectoryPath
Integer numberCapture = 1
String BulkUpload = GlobalVariable.BulkUpload
String ExcelFilename = GlobalVariable.ExcelFilename
String stepName = GlobalVariable.stepName

TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, numberCapture++, 'Pilih Use Case '+ UseCase, newDirectoryPath, true, false)
WebUI.click(findTestObject('Object Repository/COP/CardActivity/div_Card Bucket Adjustment'))

// Path ke file Excel
String excelFilePath = RunConfiguration.getProjectDir() + GlobalVariable.PathDataExcel
FileInputStream file = new FileInputStream(excelFilePath)
Workbook workbook = new XSSFWorkbook(file)
Sheet sheetAct = workbook.getSheet("Act Bucket Adjusment")

if (BulkUpload == 'Y') {
	// Input Excel
	WebUI.click(findTestObject('Object Repository/COP/input_Excel'))
	TestObject uploadExel = findTestObject('Object Repository/COP/label_Upload Excel Activity')
	WebUI.uploadFile(uploadExel, ExcelFilename)
}
else {
	// Cari berdasarkan TC
	String NoRek = ""
	String Pokok = ""
	String Bunga = ""
	String Biaya = ""
	String Denda = ""
	String Teoritis = ""
	for (int i = 2; i <= sheetAct.getLastRowNum(); i++) {
		Row row = sheetAct.getRow(i)
		if (row != null && ExcelHelper.getCellValueAsString(row, 0) == NoTC && ExcelHelper.getCellValueAsString(row, 1) == Seq) {
			NoRek = ExcelHelper.getCellValueAsString(row, 4)
			Pokok = ExcelHelper.getCellValueAsString(row, 5)
			Bunga = ExcelHelper.getCellValueAsString(row, 6)
			Biaya = ExcelHelper.getCellValueAsString(row, 7)
			Denda = ExcelHelper.getCellValueAsString(row, 8)
			Teoritis = ExcelHelper.getCellValueAsString(row, 9)
			break
		}
	}
	WebDriver driver = DriverFactory.getWebDriver()
	WebUI.comment("TC: ${NoTC}")
	
	println("NoRek: "+NoRek)
	WebUI.setText(findTestObject('Object Repository/Activity/ActivityBucketAdjusment_Object/input_NoRekening'), NoRek)
	println("Pokok: "+Pokok)
	WebUI.setText(findTestObject('Object Repository/Activity/ActivityBucketAdjusment_Object/input_Pokok'), Pokok != null ? Pokok : '')
	println("Bunga: "+Bunga)
	WebUI.setText(findTestObject('Object Repository/Activity/ActivityBucketAdjusment_Object/input_Bunga'), Bunga != null ? Bunga : '')
	println("Biaya: "+Biaya)
	WebUI.setText(findTestObject('Object Repository/Activity/ActivityBucketAdjusment_Object/input_Biaya'), Biaya != null ? Biaya : '')
	println("Denda: "+Denda)
	WebUI.setText(findTestObject('Object Repository/Activity/ActivityBucketAdjusment_Object/input_Denda'), Denda != null ? Denda : '')
	println("Teoritis: "+Teoritis)
	WebUI.setText(findTestObject('Object Repository/Activity/ActivityBucketAdjusment_Object/input_Teoritis'), Teoritis != null ? Teoritis : '')
}

ActivityUtils.saveActivityAndCapture(NoTC, stepName, newDirectoryPath, numberCapture)
