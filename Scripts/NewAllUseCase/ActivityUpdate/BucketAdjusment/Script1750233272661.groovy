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
// Path ke file Excel
String excelFilePath = RunConfiguration.getProjectDir() + GlobalVariable.PathDataExcel
FileInputStream file = new FileInputStream(excelFilePath)
Workbook workbook = new XSSFWorkbook(file)
Sheet sheetAct = workbook.getSheet("Act Bucket Adjusment")

List<String> ArrayWorstColForPokok = ['01', '02', '03', '04', '05']
List<String> ArrayWorstColForBiaya = ['01', '02']

println("Seq: "+Seq)
// Cari berdasarkan TC
String RekSingleSidePokok = ""
String RekSingleSideBiaya = ""
for (int i = 2; i <= sheetAct.getLastRowNum(); i++) {
	Row row = sheetAct.getRow(i)
	if (row != null && ExcelHelper.getCellValueAsString(row, 0) == NoTC && ExcelHelper.getCellValueAsString(row, 1) == Seq) {
//	if (row != null && ExcelHelper.getCellValueAsString(row, 0) == NoTC) {
		RekSingleSidePokok = ExcelHelper.getCellValueAsString(row, 10)
		RekSingleSideBiaya = ExcelHelper.getCellValueAsString(row, 11)
		break
	}
}
println("RekSingleSidePokok: "+RekSingleSidePokok)
println("RekSingleSideBiaya: "+RekSingleSideBiaya)

String worstCollectibility = WebUI.getAttribute(findTestObject('Object Repository/Activity/ActivityBucketAdjusment_Object/input_WorstCollectibility'), 'value')
println("WorstCollectibility value: " + worstCollectibility)
String KodeStruk = WebUI.getAttribute(findTestObject('Object Repository/Activity/ActivityBucketAdjusment_Object/input_KodeStruk'), 'value')
println("KodeStruk value: " + KodeStruk)

if (ArrayWorstColForPokok.contains(worstCollectibility.substring(0, 2))) {
	println("input NoRekSingleSide Pokok")
	WebUI.setText(findTestObject('Object Repository/Activity/ActivityBucketAdjusment_Object/input_NorekSingleSide_pokok'), RekSingleSidePokok)
}
if (ArrayWorstColForBiaya.contains(worstCollectibility.substring(0, 2))) {
	if (KodeStruk == '') {		
		println("input NoRekSingleSide Biaya")
		WebUI.setText(findTestObject('Object Repository/Activity/ActivityBucketAdjusment_Object/input_NorekSingleSide_biaya'), RekSingleSideBiaya)
	}
	else {
		if (KodeStruk.substring(0,1) == '8') {	
			println("input NoRekSingleSide Biaya")
			WebUI.setText(findTestObject('Object Repository/Activity/ActivityBucketAdjusment_Object/input_NorekSingleSide_biaya'), RekSingleSideBiaya)
		}
	}
}