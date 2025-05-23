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
import java.text.SimpleDateFormat
import java.util.Date
import excel.ExcelHelper

String testCaseName = RunConfiguration.getExecutionSourceName()
String NoTC = GlobalVariable.NoTC
String Pencairan = GlobalVariable.Pencairan
String Skenario = GlobalVariable.Skenario
//String Nominal = GlobalVariable.Nominal
String newDirectoryPath = GlobalVariable.newDirectoryPath
Integer numberCapture = 1

// Path ke file Excel
String excelFilePath = RunConfiguration.getProjectDir() + "/Data Files/NewSkenarioAllActivity.xlsx"
FileInputStream file = new FileInputStream(excelFilePath)
Workbook workbook = new XSSFWorkbook(file)
Sheet sheet3 = workbook.getSheet("BurekPenc")

String Cif = ""
String RekKredit = ""
String Nominal = ""
for (int i = 1; i <= sheet3.getLastRowNum(); i++) {
	Row row = sheet3.getRow(i)
	if (row != null && row.getCell(0).getStringCellValue() == NoTC) {
//		Cif = String.valueOf((long) row.getCell(5).getNumericCellValue())
		Cif = ExcelHelper.getCellValueAsString(row, 5)
//		RekKredit = String.valueOf((long) row.getCell(6).getNumericCellValue())
		RekKredit = ExcelHelper.getCellValueAsString(row, 6)
		Nominal = String.valueOf((long) row.getCell(7).getNumericCellValue())
		break
	}
}

WebUI.comment("TC: ${NoTC}; Pencairan: ${Pencairan} ; Cif: ${Cif}; Nominal ${Nominal} RekKredit: ${RekKredit}")

WebUI.click(findTestObject('Object Repository/COP/TabCard/a_Tab_AsuransiPembukaan'))
WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Card Activity.png')
WebUI.click(findTestObject('Object Repository/COP/CardActivity/div_Card_Pembukaan Rekening Pinjaman'))

WebUI.setText(findTestObject('Object Repository/Activity/ActivityPembukaanRek_Object/input_CIF'), Cif)
WebUI.click(findTestObject('Object Repository/Activity/ActivityPembukaanRek_Object/a_Search Cif'))
WebUI.delay(3)
//WebUI.click(findTestObject('Object Repository/Activity/ActivityPembukaanRek_Object/button_OK_Alert Form Pembukaan'))

WebDriver driver = DriverFactory.getWebDriver()
WebUI.click(findTestObject('Object Repository/Activity/ActivityPembukaanRek_Object/span_Cabang Pembuka'))
WebUI.delay(1)
List<WebElement> subCabPembukaOptions = driver.findElements(By.xpath("//li[contains(@class,'select2-results__option') and contains(text(),'760 : DIVISI OPERASIONAL - JPC')]"))
subCabPembukaOptions[0].click()
WebUI.click(findTestObject('Object Repository/Activity/ActivityPembukaanRek_Object/span_Sub Kategori'))
WebUI.delay(1)
List<WebElement> subKategoriOptions = driver.findElements(By.xpath("//li[contains(@class,'select2-results__option') and contains(text(),'Efektif IDR')]"))
subKategoriOptions[0].click()

WebUI.setText(findTestObject('Object Repository/Activity/ActivityPembukaanRek_Object/input_Maksimum Kredit'), Nominal.replaceAll("[^0-9]", ""))

if (Pencairan == 'Pencairan Pertama') {
	WebUI.click(findTestObject('Object Repository/Activity/ActivityPembukaanRek_Object/input_Pencairan Dana_Pencarian Pertama'))
}
else {
	WebUI.click(findTestObject('Object Repository/Activity/ActivityPembukaanRek_Object/input_Pencairan Dana_Pencairan Kedua'))
}

WebUI.setText(findTestObject('Object Repository/Activity/ActivityPembukaanRek_Object/input_Nominal_Pencairan Dana'), Nominal.replaceAll("[^0-9]", ""))
WebUI.setText(findTestObject('Object Repository/Activity/ActivityPembukaanRek_Object/input_No. Rekening Simp_PencairanDana'), RekKredit)
// Ambil tanggal dan waktu sekarang
Date now = new Date()
SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss")
String DataTimeNow = sdf.format(now)
WebUI.setText(findTestObject('Object Repository/Activity/ActivityPembukaanRek_Object/textarea_Narasi PencairanDana'), "ASIC-"+DataTimeNow+"-"+NoTC)
WebUI.setText(findTestObject('Object Repository/Activity/ActivityPembukaanRek_Object/textarea_Narasi Tambahan_PencairanDana'), 'Test narasi 2')

CustomKeywords.'custom.CustomKeywords.captureFullPageInSections'(newDirectoryPath+'/', numberCapture++ +'. form')

WebUI.click(findTestObject('Object Repository/Activity/ActivityPembukaanRek_Object/button_Save_Form_Pembukaan'))
WebUI.waitForElementVisible(findTestObject('Object Repository/Activity/ActivityPembukaanRek_Object/button_OK_Alert Form Pembukaan'), 30)
WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Sukses Save.png')
WebUI.click(findTestObject('Object Repository/Activity/ActivityPembukaanRek_Object/button_OK_Alert Form Pembukaan'))