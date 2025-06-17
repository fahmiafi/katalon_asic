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

WebUI.click(findTestObject('Object Repository/COP/TabCard/a_Tab_PenutupanPendingRestrukturisasi'))
TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, numberCapture++, 'Pilih Use Case '+ UseCase, newDirectoryPath, true, false)
WebUI.click(findTestObject('Object Repository/COP/CardActivity/div_Card PenutupanRek'))

// Path ke file Excel
String excelFilePath = RunConfiguration.getProjectDir() + GlobalVariable.PathDataExcel
FileInputStream file = new FileInputStream(excelFilePath)
Workbook workbook = new XSSFWorkbook(file)
Sheet sheetAct = workbook.getSheet("Act Penutupan Rek")

if (BulkUpload == 'Y') {
	// Input Excel
	WebUI.click(findTestObject('Object Repository/COP/input_Excel'))
	TestObject uploadExel = findTestObject('Object Repository/COP/label_Upload Excel Activity')
	WebUI.uploadFile(uploadExel, ExcelFilename)
}
else {
	// Cari berdasarkan TC
	String NoRek = ""
	String BiayaAdminPSJT = ""
	String BebasBiayaTutupRekening = ""
	String NominalOverride = ""
	String RekPembebananSaldoPinjaman = ""
	String RekPembebananBiayaLainnya = ""
	String Narasi = ""
	String NarasiTambahan = ""
	for (int i = 2; i <= sheetAct.getLastRowNum(); i++) {
		Row row = sheetAct.getRow(i)
		if (row != null && ExcelHelper.getCellValueAsString(row, 0) == NoTC && ExcelHelper.getCellValueAsString(row, 1) == Seq) {
			NoRek 						= ExcelHelper.getCellValueAsString(row, 4)
			BiayaAdminPSJT 				= ExcelHelper.getCellValueAsString(row, 5)
			BebasBiayaTutupRekening 	= ExcelHelper.getCellValueAsString(row, 6)
			NominalOverride 			= ExcelHelper.getCellValueAsString(row, 7)
			RekPembebananSaldoPinjaman 	= ExcelHelper.getCellValueAsString(row, 8)
			RekPembebananBiayaLainnya 	= ExcelHelper.getCellValueAsString(row, 9)
			Narasi 						= ExcelHelper.getCellValueAsString(row, 10)
			NarasiTambahan 				= ExcelHelper.getCellValueAsString(row, 11)
			break
		}
	}
	WebDriver driver = DriverFactory.getWebDriver()
	WebUI.comment("TC: ${NoTC}")
	
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
}

ActivityUtils.saveActivityAndCapture(NoTC, stepName, newDirectoryPath, numberCapture)
