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
String Pencairan = GlobalVariable.Pencairan
String newDirectoryPath = GlobalVariable.newDirectoryPath
Integer numberCapture = 1
List<String> PeriodeKapitalisasiList = [
	'001: Bulanan', 
	'002: Dua Bulanan', 
	'003: Tiga Bulanan', 
	'004: Empat Bulanan', 
	'006: Semesteran', 
	'012: Tahunan', 
]

// Path ke file Excel
String excelFilePath = RunConfiguration.getProjectDir() + "/Data Files/NewSkenarioAllActivity.xlsx"
FileInputStream file = new FileInputStream(excelFilePath)
Workbook workbook = new XSSFWorkbook(file)
Sheet sheet3 = workbook.getSheet("MaintenanceRek")

// Cari berdasarkan TC
String NoRek = ""
String PosisiRek_RekeningAfiliasi = ""
String PosisiRek_RateBunga = ""
String KapitalisasiBunga_PeriodeKapitalisasi = ""
String KapitalisasiBunga_TanggalMulaiKapitalisasi = ""
String KapitalisasiBunga_HariPembayaran = ""
String KapitalisasiBunga_TanggalKapitalisasi = ""
String Maksimum_Rekening = ""
for (int i = 1; i <= sheet3.getLastRowNum(); i++) {
	Row row = sheet3.getRow(i)
	if (row != null && row.getCell(0).getStringCellValue() == NoTC) {
		NoRek = String.valueOf((long) row.getCell(5).getNumericCellValue())
//		PosisiRek_RekeningAfiliasi = String.valueOf((long) row.getCell(6).getNumericCellValue())
//		PosisiRek_RateBunga = String.valueOf((long) row.getCell(7).getNumericCellValue())
//		KapitalisasiBunga_PeriodeKapitalisasi = row.getCell(8).getStringCellValue()
//		KapitalisasiBunga_TanggalMulaiKapitalisasi = row.getCell(9).getStringCellValue()
//		KapitalisasiBunga_HariPembayaran = String.valueOf((long) row.getCell(10).getNumericCellValue())
//		KapitalisasiBunga_TanggalKapitalisasi = String.valueOf((long) row.getCell(11).getNumericCellValue())
//		Maksimum_Rekening = String.valueOf((long) row.getCell(12).getNumericCellValue())
		
//		NoRek = ExcelHelper.getCellValueAsString(row, 5)
		PosisiRek_RekeningAfiliasi = ExcelHelper.getCellValueAsString(row, 6)
		PosisiRek_RateBunga = ExcelHelper.getCellValueAsString(row, 7)
		KapitalisasiBunga_PeriodeKapitalisasi = ExcelHelper.getCellValueAsString(row, 8)
		KapitalisasiBunga_TanggalMulaiKapitalisasi = ExcelHelper.getCellValueAsString(row, 9)
		KapitalisasiBunga_HariPembayaran = ExcelHelper.getCellValueAsString(row, 10)
		KapitalisasiBunga_TanggalKapitalisasi = ExcelHelper.getCellValueAsString(row, 11)
		Maksimum_Rekening = ExcelHelper.getCellValueAsString(row, 12)
		break
	}
}
WebDriver driver = DriverFactory.getWebDriver()
WebUI.comment("TC: ${NoTC}")

WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Card Activity.png')
WebUI.click(findTestObject('Object Repository/COP/TabCard/a_ Tab_Maintenance Rekening'))
WebUI.click(findTestObject('Object Repository/COP/CardActivity/div_Card Maintenance Rekening Pinjaman'))
WebUI.setText(findTestObject('Object Repository/Activity/ActivityMaintenanceRek_Object/input_Nomor Rekening'), NoRek)

Boolean BoolPosisiRek = false
if (PosisiRek_RekeningAfiliasi != null) {
	BoolPosisiRek = true
	println("PosisiRek akan dijalankan")
}
Boolean BoolKapitalisasiBunga = false
if (KapitalisasiBunga_PeriodeKapitalisasi != null || KapitalisasiBunga_TanggalMulaiKapitalisasi != null ||
	KapitalisasiBunga_HariPembayaran != null || KapitalisasiBunga_TanggalKapitalisasi != null) {
	BoolKapitalisasiBunga = true
	println("KapitalisasiBunga akan dijalankan")
}
Boolean BoolMaksimum_Rekening = false
if (Maksimum_Rekening != null) {
	BoolMaksimum_Rekening = true
	println("Maksimum_Rekening akan dijalankan")
}


if (BoolPosisiRek) {
	println ("PosisiRek_RekeningAfiliasi : "+PosisiRek_RekeningAfiliasi)
	WebUI.check(findTestObject('Object Repository/Activity/ActivityMaintenanceRek_Object/check_PosisiRek_RekeningAfiliasi'))
	WebUI.setText(findTestObject('Object Repository/Activity/ActivityMaintenanceRek_Object/input_PosisiRek_Rekening Afiliasi_Menjadi'), PosisiRek_RekeningAfiliasi)
}

//println ("PosisiRek_RateBunga : "+PosisiRek_RateBunga)
//WebUI.check(findTestObject('Object Repository/Activity/ActivityMaintenanceRek_Object/check_PosisiRek_RateBunga'))
//WebUI.setText(findTestObject('Object Repository/Activity/ActivityMaintenanceRek_Object/input_PosisiRek_RateBunga_Menjadi'), PosisiRek_RateBunga)

if (BoolKapitalisasiBunga) {
	println ("KapitalisasiBunga_PeriodeKapitalisasi : "+KapitalisasiBunga_PeriodeKapitalisasi)
	WebUI.check(findTestObject('Object Repository/Activity/ActivityMaintenanceRek_Object/check_KapitalisasiBunga'))
	String resultPeriodeKapitalisasi = PeriodeKapitalisasiList.find { it.contains(KapitalisasiBunga_PeriodeKapitalisasi) }
	WebUI.selectOptionByLabel(findTestObject('Object Repository/Activity/ActivityMaintenanceRek_Object/select_PeriodeKapitalisasi_Menjadi'), resultPeriodeKapitalisasi, true)
	
	println ("KapitalisasiBunga_TanggalMulaiKapitalisasi : "+KapitalisasiBunga_TanggalMulaiKapitalisasi)
	WebUI.setText(findTestObject('Object Repository/Activity/ActivityMaintenanceRek_Object/input_PeriodeKapitalisasi_TanggalMulaiKapitalisasi_Menjadi'), KapitalisasiBunga_TanggalMulaiKapitalisasi)
	
	println ("KapitalisasiBunga_HariPembayaran : "+KapitalisasiBunga_HariPembayaran)
	WebUI.setText(findTestObject('Object Repository/Activity/ActivityMaintenanceRek_Object/input_PeriodeKapitalisasi_HariPembayaran_Menjadi'), KapitalisasiBunga_HariPembayaran)
	
	println ("KapitalisasiBunga_TanggalKapitalisasi : "+KapitalisasiBunga_TanggalKapitalisasi)
	WebUI.setText(findTestObject('Object Repository/Activity/ActivityMaintenanceRek_Object/input_PeriodeKapitalisasi_TanggalKapitalisasi_Menjadi'), KapitalisasiBunga_TanggalKapitalisasi)
}

if (BoolMaksimum_Rekening) {
	println ("Maksimum_Rekening : "+Maksimum_Rekening.trim().length())
	WebUI.check(findTestObject('Object Repository/Activity/ActivityMaintenanceRek_Object/check_MaksimumPinjaman'))
	WebUI.setText(findTestObject('Object Repository/Activity/ActivityMaintenanceRek_Object/input_Maksimum Rekening'), Maksimum_Rekening)
}

CustomKeywords.'custom.CustomKeywords.captureFullPageInSections'(newDirectoryPath+'/', numberCapture++ +'. form')

WebUI.click(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/button_Save'))
WebUI.waitForElementVisible(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/button_Save OK'), 30)
WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Sukses Save.png')
WebUI.click(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/button_Save OK'))