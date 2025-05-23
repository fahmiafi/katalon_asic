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
String excelFilePath = RunConfiguration.getProjectDir() + "/Data Files/NewSkenarioAllActivity.xlsx"
FileInputStream file = new FileInputStream(excelFilePath)
Workbook workbook = new XSSFWorkbook(file)
Sheet sheet3 = workbook.getSheet("BlokirRek")

// Cari berdasarkan TC
String NoRek = ""
String IsPasang = ""
String Nominal = ""
for (int i = 1; i <= sheet3.getLastRowNum(); i++) {
	Row row = sheet3.getRow(i)
	if (row != null && row.getCell(0).getStringCellValue() == NoTC) {
		NoRek = String.valueOf((long) row.getCell(5).getNumericCellValue())
		IsPasang = row.getCell(6).getStringCellValue()
		Nominal = String.valueOf((long) row.getCell(7).getNumericCellValue())
		break
	}
}
WebDriver driver = DriverFactory.getWebDriver()
WebUI.comment("TC: ${NoTC}")

WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Card Activity.png')
WebUI.click(findTestObject('Object Repository/COP/CardActivity/div_Card BlokirRek'))
WebUI.setText(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/input_NomorRekening'), NoRek)
String valPasang = "Pasang"
if (IsPasang != "Pasang") {
	valPasang = "Lepas"
}
WebUI.selectOptionByValue(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/select_PasangOrLepas'), valPasang, true)
if (valPasang == 'Pasang') {
	WebUI.click(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/span_PerlakuanBlokir'))
	WebUI.delay(1)
	List<WebElement> subPerlakuakBlokir = driver.findElements(By.xpath("//li[contains(@class,'select2-results__option') and contains(text(),'09: Lainnya')]"))
	subPerlakuakBlokir[0].click()	
}
WebUI.setText(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/input_NominalBlokir'), Nominal.replaceAll("[^0-9]", ""))
WebUI.setText(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/input_NarasiBlokir1'), 'Test Narasi 1')
//WebUI.setText(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/input_NarasiBlokir2'), 'Test Narasi 1')

CustomKeywords.'custom.CustomKeywords.captureFullPageInSections'(newDirectoryPath+'/', numberCapture++ +'. form')

WebUI.click(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/button_Save'))
WebUI.waitForElementVisible(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/button_Save OK'), 30)
WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Sukses Save.png')
WebUI.click(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/button_Save OK'))
