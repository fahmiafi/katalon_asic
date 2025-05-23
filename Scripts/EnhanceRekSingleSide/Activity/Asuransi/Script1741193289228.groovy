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
Sheet sheet3 = workbook.getSheet("Asuransi")

// Cari berdasarkan TC
String NomorRekening = ""
String JenisAsuransi = ""
String NominalCover = ""
String NominalPremi = ""
String ImbalJasa = ""
String BiayaPolis = ""
String BiayaMaterai = ""
String Keterangan = ""
String RekPerusahaanAsuransi = ""
String NoPolis = ""
String TanggalMulai = ""
String TanggalJatuhTempo = ""
String RekPembebananBiaya = ""
for (int i = 1; i <= sheet3.getLastRowNum(); i++) {
	Row row = sheet3.getRow(i)
	if (row != null && row.getCell(0).getStringCellValue() == NoTC) {
		NomorRekening = String.valueOf((long) row.getCell(5).getNumericCellValue())
		JenisAsuransi = row.getCell(6).getStringCellValue()
		NominalCover = String.valueOf((long) row.getCell(7).getNumericCellValue())
		NominalPremi = String.valueOf((long) row.getCell(8).getNumericCellValue())
		ImbalJasa = row.getCell(9).getStringCellValue()
		BiayaPolis = String.valueOf((long) row.getCell(10).getNumericCellValue())
		BiayaMaterai = String.valueOf((long) row.getCell(11).getNumericCellValue())
		Keterangan = row.getCell(12).getStringCellValue()
		RekPerusahaanAsuransi = String.valueOf((long) row.getCell(13).getNumericCellValue())
		NoPolis = row.getCell(14).getStringCellValue()
		TanggalMulai = row.getCell(15).getStringCellValue()
		TanggalJatuhTempo = row.getCell(16).getStringCellValue()
		RekPembebananBiaya = String.valueOf((long) row.getCell(17).getNumericCellValue())
		break
	}
}
WebDriver driver = DriverFactory.getWebDriver()
WebUI.comment("TC: ${NoTC}")

println ("NomorRekening: "+NomorRekening)
WebUI.setText(findTestObject('Object Repository/Activity/ActivityAsuransi_Object/input_Nomor Rekening'), NomorRekening)

println ("JenisAsuransi: "+JenisAsuransi)
WebUI.selectOptionByLabel(findTestObject('Object Repository/Activity/ActivityAsuransi_Object/select_Jenis Asuransi'), JenisAsuransi, true)

println ("NominalCover: "+NominalCover)
WebUI.setText(findTestObject('Object Repository/Activity/ActivityAsuransi_Object/input_Nominal Cover'), NominalCover.replaceAll("[^0-9]", ""))

println ("NominalPremi: "+NominalPremi)
WebUI.setText(findTestObject('Object Repository/Activity/ActivityAsuransi_Object/input_Nominal Premi'), NominalPremi.replaceAll("[^0-9]", ""))

println ("ImbalJasa: "+ImbalJasa)
WebUI.selectOptionByLabel(findTestObject('Object Repository/Activity/ActivityAsuransi_Object/select_Imbal Jasa'), ImbalJasa, true)

println ("BiayaPolis: "+BiayaPolis)
WebUI.setText(findTestObject('Object Repository/Activity/ActivityAsuransi_Object/input_Biaya Polis'), BiayaPolis.replaceAll("[^0-9]", ""))

println ("BiayaMaterai: "+BiayaMaterai)
WebUI.setText(findTestObject('Object Repository/Activity/ActivityAsuransi_Object/input_Biaya Materai'), BiayaPolis.replaceAll("[^0-9]", ""))

println ("Keterangan: "+Keterangan)
WebUI.setText(findTestObject('Object Repository/Activity/ActivityAsuransi_Object/textarea_Keterangan'), Keterangan)

println ("RekPerusahaanAsuransi: "+RekPerusahaanAsuransi)
WebUI.setText(findTestObject('Object Repository/Activity/ActivityAsuransi_Object/input_Rek. Perusahaan Asuransi'), RekPerusahaanAsuransi)

println ("NoPolis: "+NoPolis)
WebUI.setText(findTestObject('Object Repository/Activity/ActivityAsuransi_Object/input_No Polis'), NoPolis)

println ("TanggalMulai: "+TanggalMulai)
WebUI.setText(findTestObject('Object Repository/Activity/ActivityAsuransi_Object/input_Tanggal Mulai'), TanggalMulai)

println ("TanggalJatuhTempo: "+TanggalJatuhTempo)
WebUI.setText(findTestObject('Object Repository/Activity/ActivityAsuransi_Object/input_Tanggal Jatuh Tempo'), TanggalJatuhTempo)

println ("RekPembebananBiaya: "+RekPembebananBiaya)
WebUI.setText(findTestObject('Object Repository/Activity/ActivityAsuransi_Object/input_Rek. Pembebanan Biaya'), RekPembebananBiaya)
WebUI.selectOptionByLabel(findTestObject('Object Repository/Activity/ActivityAsuransi_Object/select_Jenis Rek Pembebanan Biaya'), 'Simpanan', true)

WebUI.click(findTestObject('Object Repository/Activity/ActivityAsuransi_Object/textarea_Keterangan'))

CustomKeywords.'custom.CustomKeywords.captureFullPageInSections'(newDirectoryPath+'/', numberCapture++ +'. form')

WebUI.click(findTestObject('Object Repository/Activity/ActivityAsuransi_Object/button_Save'))
WebUI.waitForElementVisible(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/button_Save OK'), 30)
WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Sukses Save.png')
WebUI.click(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/button_Save OK'))