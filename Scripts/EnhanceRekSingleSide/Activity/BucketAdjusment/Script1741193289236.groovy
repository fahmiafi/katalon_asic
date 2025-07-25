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

String testCaseName = RunConfiguration.getExecutionSourceName()
String NoTC = GlobalVariable.NoTC
String Pencairan = GlobalVariable.Pencairan
String newDirectoryPath = GlobalVariable.newDirectoryPath
Integer numberCapture = 1

// Path ke file Excel
String excelFilePath = RunConfiguration.getProjectDir() + "/Data Files/SkenarioEnhanceRekSingleSide.xlsx"
FileInputStream file = new FileInputStream(excelFilePath)
Workbook workbook = new XSSFWorkbook(file)
Sheet sheet3 = workbook.getSheet("BucketAdjusment")

// Cari berdasarkan TC
String NoRek = ""
String RekSingleSide = ""
String Pokok = ""
String Bunga = ""
String Biaya = ""
String Denda = ""
String Teoritis = ""
for (int i = 1; i <= sheet3.getLastRowNum(); i++) {
	Row row = sheet3.getRow(i)
	if (row != null && row.getCell(0).getStringCellValue() == NoTC) {
		NoRek = String.valueOf((long) row.getCell(5).getNumericCellValue())
		RekSingleSide = String.valueOf((long) row.getCell(6).getNumericCellValue())
		Pokok = String.valueOf((long) row.getCell(7).getNumericCellValue())
		Bunga = String.valueOf((long) row.getCell(8).getNumericCellValue())
		Biaya = String.valueOf((long) row.getCell(9).getNumericCellValue())
		Denda = String.valueOf((long) row.getCell(10).getNumericCellValue())
		Teoritis = String.valueOf((long) row.getCell(11).getNumericCellValue())
		break
	}
}
WebDriver driver = DriverFactory.getWebDriver()
WebUI.comment("TC: ${NoTC}")

println("NoRek: "+NoRek)
WebUI.setText(findTestObject('Object Repository/Activity/ActivityBucketAdjusment_Object/input_NoRekening'), NoRek)
println("RekSingleSide: "+RekSingleSide)
WebUI.setText(findTestObject('Object Repository/Activity/ActivityBucketAdjusment_Object/input_NorekSingleSide_backup'), RekSingleSide)
println("Pokok: "+Pokok)
WebUI.setText(findTestObject('Object Repository/Activity/ActivityBucketAdjusment_Object/input_Pokok'), Pokok)
println("Bunga: "+Bunga)
WebUI.setText(findTestObject('Object Repository/Activity/ActivityBucketAdjusment_Object/input_Bunga'), Bunga)
println("Biaya: "+Biaya)
WebUI.setText(findTestObject('Object Repository/Activity/ActivityBucketAdjusment_Object/input_Biaya'), Biaya)
println("Denda: "+Denda)
WebUI.setText(findTestObject('Object Repository/Activity/ActivityBucketAdjusment_Object/input_Denda'), Denda)
println("Teoritis: "+Teoritis)
WebUI.setText(findTestObject('Object Repository/Activity/ActivityBucketAdjusment_Object/input_Teoritis'), Teoritis)

CustomKeywords.'custom.CustomKeywords.captureFullPageInSections'(newDirectoryPath+'/', numberCapture++ +'. form')

WebUI.click(findTestObject('Object Repository/Activity/ActivityBucketAdjusment_Object/button_Save'))
WebUI.waitForElementVisible(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/button_Save OK'), 30)
WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Sukses Save.png')
WebUI.click(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/button_Save OK'))
