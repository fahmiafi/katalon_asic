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
Sheet sheet3 = workbook.getSheet("PenutupanRek")

// Cari berdasarkan TC
String NoRek = ""
String BiayaAdminPSJT = ""
String BebasBiayaTutupRekening = ""
String NominalOverride = ""
String RekPembebananSaldoPinjaman = ""
String RekPembebananBiayaLainnya = ""
String Narasi = ""
String NarasiTambahan = ""
for (int i = 1; i <= sheet3.getLastRowNum(); i++) {
	Row row = sheet3.getRow(i)
	if (row != null && row.getCell(0).getStringCellValue() == NoTC) {
		NoRek = String.valueOf((long) row.getCell(5).getNumericCellValue())
		BiayaAdminPSJT = row.getCell(6).getStringCellValue()
		BebasBiayaTutupRekening = row.getCell(7).getStringCellValue()
		NominalOverride = String.valueOf((long) row.getCell(8).getNumericCellValue())
		RekPembebananSaldoPinjaman = String.valueOf((long) row.getCell(9).getNumericCellValue())
		RekPembebananBiayaLainnya = String.valueOf((long) row.getCell(10).getNumericCellValue())
		Narasi = row.getCell(11).getStringCellValue()
		NarasiTambahan = row.getCell(12).getStringCellValue()
		break
	}
}
WebDriver driver = DriverFactory.getWebDriver()
WebUI.comment("TC: ${NoTC}")

WebUI.click(findTestObject('Object Repository/COP/TabCard/a_Tab_PenutupanPendingRestrukturisasi'))
WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Card Activity.png')
WebUI.click(findTestObject('Object Repository/COP/CardActivity/div_Card PenutupanRek'))

println ("NoRek: "+NoRek)
WebUI.setText(findTestObject('Object Repository/Activity/ActivityPenutupanRek_Object/input_NoRekeningPinjaman'), NoRek)

println ("BiayaAdminPSJT: "+BiayaAdminPSJT)
WebUI.selectOptionByValue(findTestObject('Object Repository/Activity/ActivityPenutupanRek_Object/select_BiayaAdminPSJT'), BiayaAdminPSJT, true)

println ("BebasBiayaTutupRekening: "+BebasBiayaTutupRekening)
WebUI.selectOptionByValue(findTestObject('Object Repository/Activity/ActivityPenutupanRek_Object/select_BebasBiayaTutupRek'), BebasBiayaTutupRekening, true)

println ("NominalOverride: "+NominalOverride)
WebUI.setText(findTestObject('Object Repository/Activity/ActivityPenutupanRek_Object/input_Nominal Override'), NominalOverride.replaceAll("[^0-9]", ""))

println ("RekPembebananSaldoPinjaman: "+RekPembebananSaldoPinjaman)
WebUI.setText(findTestObject('Object Repository/Activity/ActivityPenutupanRek_Object/input_RekeningPembebananSaldoPinjaman_Debet1'), RekPembebananSaldoPinjaman)

println ("RekPembebananBiayaLainnya: "+RekPembebananBiayaLainnya)
WebUI.setText(findTestObject('Object Repository/Activity/ActivityPenutupanRek_Object/input_RekeningPembebananBiayaLainnya_Debet2'), RekPembebananBiayaLainnya)

println ("Narasi: "+Narasi)
WebUI.setText(findTestObject('Object Repository/Activity/ActivityPenutupanRek_Object/input_Narasi1'), Narasi)

println ("NarasiTambahan: "+NarasiTambahan)
WebUI.setText(findTestObject('Object Repository/Activity/ActivityPenutupanRek_Object/input_Narasi2_Tambahan'), NarasiTambahan)

CustomKeywords.'custom.CustomKeywords.captureFullPageInSections'(newDirectoryPath+'/', numberCapture++ +'. form')

WebUI.click(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/button_Save'))
WebUI.waitForElementVisible(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/button_Save OK'), 30)
WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Sukses Save.png')
WebUI.click(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/button_Save OK'))
