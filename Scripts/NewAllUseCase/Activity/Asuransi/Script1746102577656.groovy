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

WebUI.click(findTestObject('Object Repository/COP/TabCard/a_Tab_AsuransiPembukaan'))
TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, numberCapture++, 'Pilih Use Case '+ UseCase, newDirectoryPath, true, false)
WebUI.click(findTestObject('Object Repository/COP/CardActivity/div_Card Asuransi Jaminan'))

// Path ke file Excel
String excelFilePath = RunConfiguration.getProjectDir() + GlobalVariable.PathDataExcel
FileInputStream file = new FileInputStream(excelFilePath)
Workbook workbook = new XSSFWorkbook(file)
Sheet sheetAct = workbook.getSheet("Act Asuransi")

if (BulkUpload == 'Y') {
	// Input Excel
	WebUI.click(findTestObject('Object Repository/COP/input_Excel'))
	TestObject uploadExel = findTestObject('Object Repository/COP/label_Upload Excel Activity')
	WebUI.uploadFile(uploadExel, ExcelFilename)
}
else {
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
	for (int i = 2; i <= sheetAct.getLastRowNum(); i++) {
		Row row = sheetAct.getRow(i)
		if (row != null && ExcelHelper.getCellValueAsString(row, 0) == NoTC && ExcelHelper.getCellValueAsString(row, 1) == Seq) {
			NomorRekening 			= ExcelHelper.getCellValueAsString(row, 4)
			JenisAsuransi 			= ExcelHelper.getCellValueAsString(row, 5)
			NominalCover 			= ExcelHelper.getCellValueAsString(row, 6)
			NominalPremi 			= ExcelHelper.getCellValueAsString(row, 7)
			ImbalJasa 				= ExcelHelper.getCellValueAsString(row, 8)
			BiayaPolis 				= ExcelHelper.getCellValueAsString(row, 9)
			BiayaMaterai 			= ExcelHelper.getCellValueAsString(row, 10)
			Keterangan 				= ExcelHelper.getCellValueAsString(row, 11)
			RekPerusahaanAsuransi 	= ExcelHelper.getCellValueAsString(row, 12)
			NoPolis 				= ExcelHelper.getCellValueAsString(row, 13)
			TanggalMulai 			= ExcelHelper.getCellValueAsString(row, 14)
			TanggalJatuhTempo 		= ExcelHelper.getCellValueAsString(row, 15)
			RekPembebananBiaya 		= ExcelHelper.getCellValueAsString(row, 16)
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
}

ActivityUtils.saveActivityAndCapture(NoTC, stepName, newDirectoryPath, numberCapture)