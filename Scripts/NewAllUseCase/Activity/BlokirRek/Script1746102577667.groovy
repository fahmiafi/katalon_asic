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

TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, numberCapture++, 3, 'Pilih Use Case '+ UseCase, newDirectoryPath, true, false)
WebUI.click(findTestObject('Object Repository/COP/CardActivity/div_Card BlokirRek'))

// Path ke file Excel
String excelFilePath = RunConfiguration.getProjectDir() + GlobalVariable.PathDataExcel
FileInputStream file = new FileInputStream(excelFilePath)
Workbook workbook = new XSSFWorkbook(file)
Sheet sheetAct = workbook.getSheet("Act Blokir Rek")

if (BulkUpload == 'Y') {
	// Input Excel
	WebUI.click(findTestObject('Object Repository/COP/input_Excel'))
	TestObject uploadExel = findTestObject('Object Repository/COP/label_Upload Excel Activity')
	WebUI.uploadFile(uploadExel, ExcelFilename)
}
else {
	// Cari berdasarkan TC
	String NoRek = ""
	String IsPasang = ""
	String Nominal = ""
	for (int i = 2; i <= sheetAct.getLastRowNum(); i++) {
		Row row = sheetAct.getRow(i)
		if (row != null && ExcelHelper.getCellValueAsString(row, 0) == NoTC && ExcelHelper.getCellValueAsString(row, 1) == Seq) {
			NoRek 		= ExcelHelper.getCellValueAsString(row, 4)
			IsPasang 	= ExcelHelper.getCellValueAsString(row, 5)
			Nominal 	= ExcelHelper.getCellValueAsString(row, 6)
			break
		}
	}
	WebDriver driver = DriverFactory.getWebDriver()
	WebUI.comment("TC: ${NoTC}")
	
	println("NoRek : "+NoRek)
	WebUI.setText(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/input_NomorRekening'), NoRek)
	String valPasang = "Pasang"
	if (IsPasang != "Pasang") {
		valPasang = "Lepas"
	}
	println("IsPasang : "+IsPasang)
	WebUI.selectOptionByValue(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/select_PasangOrLepas'), valPasang, true)
	if (valPasang == 'Pasang') {
		WebUI.click(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/span_PerlakuanBlokir'))
		WebUI.delay(1)
		List<WebElement> subPerlakuakBlokir = driver.findElements(By.xpath("//li[contains(@class,'select2-results__option') and contains(text(),'09: Lainnya')]"))
		subPerlakuakBlokir[0].click()	
	}
	println("Nominal : "+Nominal)
	WebUI.setText(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/input_NominalBlokir'), Nominal.replaceAll("[^0-9]", ""))
	WebUI.setText(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/input_NarasiBlokir1'), 'Test Narasi 1')
	//WebUI.setText(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/input_NarasiBlokir2'), 'Test Narasi 1')
}

ActivityUtils.saveActivityAndCapture(NoTC, stepName, newDirectoryPath, numberCapture)
