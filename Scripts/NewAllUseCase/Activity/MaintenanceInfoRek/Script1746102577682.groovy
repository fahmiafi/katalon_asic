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

String testCaseName = RunConfiguration.getExecutionSourceName()
String NoTC = GlobalVariable.NoTC
String Seq = GlobalVariable.Seq
String Pencairan = GlobalVariable.Pencairan
String newDirectoryPath = GlobalVariable.newDirectoryPath
Integer numberCapture = 1

// Path ke file Excel
String excelFilePath = RunConfiguration.getProjectDir() + GlobalVariable.PathDataExcel
FileInputStream file = new FileInputStream(excelFilePath)
Workbook workbook = new XSSFWorkbook(file)
Sheet sheetAct = workbook.getSheet("Act Maintenance Info")
// Cari berdasarkan TC
String NoRek = ""
String KodePeruntukan = ""
String SektorEkonomi = ""
String FlagKmk = ""
String NomorPKAwal = ""
String NomorPKAkhir = ""
String TanggalPKAwal = ""
String TanggalPKAkhir = ""
for (int i = 2; i <= sheetAct.getLastRowNum(); i++) {
	Row row = sheetAct.getRow(i)
	if (row != null && ExcelHelper.getCellValueAsString(row, 0) == NoTC && ExcelHelper.getCellValueAsString(row, 1) == Seq) {
		NoRek 			= ExcelHelper.getCellValueAsString(row, 4)
		KodePeruntukan 	= ExcelHelper.getCellValueAsString(row, 5)
		SektorEkonomi 	= ExcelHelper.getCellValueAsString(row, 6)
		FlagKmk 		= ExcelHelper.getCellValueAsString(row, 7)
		NomorPKAwal		= ExcelHelper.getCellValueAsString(row, 8)
		NomorPKAkhir	= ExcelHelper.getCellValueAsString(row, 9)
		TanggalPKAwal	= ExcelHelper.getCellValueAsString(row, 10)
		TanggalPKAkhir	= ExcelHelper.getCellValueAsString(row, 11)
		break
	}
}
WebDriver driver = DriverFactory.getWebDriver()
WebUI.comment("TC: ${NoTC}")

WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Card Activity.png')
WebUI.click(findTestObject('Object Repository/COP/TabCard/a_ Tab_Maintenance Rekening'))
WebUI.click(findTestObject('Object Repository/COP/CardActivity/div_Card Maintenance Informasi Rekening Pinjaman'))
WebUI.setText(findTestObject('Object Repository/Activity/ActivityMaintInfoRek_Object/input_Nomor Rekening'), NoRek)

if (KodePeruntukan != null) {	
	println("Kode Peruntukan : "+KodePeruntukan)
	WebUI.click(findTestObject('Object Repository/Activity/ActivityMaintInfoRek_Object/select2_KodePeruntukan'))
	WebUI.delay(1)
	List<WebElement> KodePeruntukanOptions = driver.findElements(By.xpath("//li[contains(@class,'select2-results__option') and contains(text(),'" + KodePeruntukan + "')]"))
	KodePeruntukanOptions[0].click()
}

if (SektorEkonomi != null) {
	println("Sektor Ekonomi : "+SektorEkonomi)
	WebUI.click(findTestObject('Object Repository/Activity/ActivityMaintInfoRek_Object/select2_SektorEkonomi'))
	WebUI.delay(1)
	List<WebElement> SektorEkonomiOptions = driver.findElements(By.xpath("//li[contains(@class,'select2-results__option') and contains(text(),'" + SektorEkonomi + "')]"))
	WebUI.delay(2)
	SektorEkonomiOptions[0].click()
}

println("NomorPKAwal :"+NomorPKAwal)
WebUI.setText(findTestObject('Object Repository/Activity/ActivityMaintInfoRek_Object/input_Nomor PK Awal Menjadi'), NomorPKAwal != null ? NomorPKAwal : '')
println("NomorPKAkhir :"+NomorPKAkhir)
WebUI.setText(findTestObject('Object Repository/Activity/ActivityMaintInfoRek_Object/input_Nomor PK Akhir Menjadi'), NomorPKAkhir != null ? NomorPKAkhir : '')
println("TanggalPKAwal :"+TanggalPKAwal)
WebUI.setText(findTestObject('Object Repository/Activity/ActivityMaintInfoRek_Object/input_Tanggal PK Awal Menjadi'), TanggalPKAwal != null ? TanggalPKAwal : '')
println("TanggalPKAkhir :"+TanggalPKAkhir)
WebUI.setText(findTestObject('Object Repository/Activity/ActivityMaintInfoRek_Object/input_Tanggal PK Akhir Menjadi'), TanggalPKAkhir != null ? TanggalPKAkhir : '')

if (FlagKmk != null) {
	println("Flag KMK : "+FlagKmk)
	WebUI.click(findTestObject('Object Repository/Activity/ActivityMaintInfoRek_Object/select2_FlagKMK'))
	WebUI.delay(1)
	List<WebElement> FlagKmkOptions = driver.findElements(By.xpath("//li[contains(@class,'select2-results__option') and contains(text(),'" + FlagKmk + "')]"))
	FlagKmkOptions[0].click()
}

CustomKeywords.'custom.CustomKeywords.scrollToTop'()
WebUI.click(findTestObject('Object Repository/Activity/ActivityMaintInfoRek_Object/input_Nomor Rekening'))

CustomKeywords.'custom.CustomKeywords.captureFullPageInSections'(newDirectoryPath+'/', numberCapture++ +'. Input Form')

WebUI.scrollToElement(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/button_Save'), 30)
WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Simpan.png')
WebUI.click(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/button_Save'))
WebUI.waitForElementVisible(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/button_Save OK'), 30)
WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Berhasil disimpan.png')
WebUI.click(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/button_Save OK'))