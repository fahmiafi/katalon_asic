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

List<String> PeriodeKapitalisasiList = [
	'001: Bulanan',
	'002: Dua Bulanan',
	'003: Tiga Bulanan',
	'004: Empat Bulanan',
	'006: Semesteran',
	'012: Tahunan',
]

WebUI.click(findTestObject('Object Repository/COP/TabCard/a_ Tab_Maintenance Rekening'))
TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, numberCapture++, 3, 'Pilih Use Case '+ UseCase, newDirectoryPath, true, false)
WebUI.click(findTestObject('Object Repository/COP/CardActivity/div_Card Maintenance Rekening Pinjaman'))

if (BulkUpload == 'Y') {
	// Input Excel
	WebUI.click(findTestObject('Object Repository/COP/input_Excel'))
	TestObject uploadExel = findTestObject('Object Repository/COP/label_Upload Excel Activity')
	WebUI.uploadFile(uploadExel, ExcelFilename)
}
else {
	// Input Form
	// Path ke file Excel
	String excelFilePath = RunConfiguration.getProjectDir() + GlobalVariable.PathDataExcel
	FileInputStream file = new FileInputStream(excelFilePath)
	Workbook workbook = new XSSFWorkbook(file)
	Sheet sheetAct = workbook.getSheet("Act Maintenance Rek")
	
	// Cari berdasarkan TC
	String NoRek = ""
	String PosisiRek_RekeningAfiliasi = ""
	String PosisiRek_RateBunga = ""
	String KapitalisasiBunga_PeriodeKapitalisasi = ""
	String KapitalisasiBunga_TanggalMulaiKapitalisasi = ""
	String KapitalisasiBunga_HariPembayaran = ""
	String KapitalisasiBunga_TanggalKapitalisasi = ""
	String Maksimum_Rekening = ""
	String JatuhTempo_NomorPKAwal = ""
	String JatuhTempo_NomorPKAkhir = ""
	String JatuhTempo_TanggalPKAwal = ""
	String JatuhTempo_TanggalPKAkhir = ""
	String JatuhTempo_PerpanjanganSementara = ""
	for (int i = 2; i <= sheetAct.getLastRowNum(); i++) {
		Row row = sheetAct.getRow(i)
		if (row != null && ExcelHelper.getCellValueAsString(row, 0) == NoTC && ExcelHelper.getCellValueAsString(row, 1) == Seq) {
			NoRek 										= ExcelHelper.getCellValueAsString(row, 4)
			PosisiRek_RekeningAfiliasi 					= ExcelHelper.getCellValueAsString(row, 5)
			PosisiRek_RateBunga 						= ExcelHelper.getCellValueAsString(row, 6)
			KapitalisasiBunga_PeriodeKapitalisasi 		= ExcelHelper.getCellValueAsString(row, 7)
			KapitalisasiBunga_TanggalMulaiKapitalisasi 	= ExcelHelper.getCellValueAsString(row, 8)
			KapitalisasiBunga_HariPembayaran 			= ExcelHelper.getCellValueAsString(row, 9)
			KapitalisasiBunga_TanggalKapitalisasi 		= ExcelHelper.getCellValueAsString(row, 10)
			Maksimum_Rekening 							= ExcelHelper.getCellValueAsString(row, 11)
			JatuhTempo_NomorPKAwal 						= ExcelHelper.getCellValueAsString(row, 12)
			JatuhTempo_NomorPKAkhir 					= ExcelHelper.getCellValueAsString(row, 13)
			JatuhTempo_TanggalPKAwal 					= ExcelHelper.getCellValueAsString(row, 14)
			JatuhTempo_TanggalPKAkhir 					= ExcelHelper.getCellValueAsString(row, 15)
			JatuhTempo_PerpanjanganSementara			= ExcelHelper.getCellValueAsString(row, 16)
			break
		}
	}
	WebDriver driver = DriverFactory.getWebDriver()
	WebUI.comment("TC: ${NoTC}")
	
	Boolean BoolPosisiRek = false
	if (PosisiRek_RekeningAfiliasi != null) {
		BoolPosisiRek = true
		println("PosisiRek akan dijalankan")
	}
	Boolean BoolRateBunga = false
	if (PosisiRek_RateBunga != null) {
		BoolRateBunga = true
		println("PosisiRek_RateBunga akan dijalankan")
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
	Boolean BoolJatuhTempo = false
	if (JatuhTempo_NomorPKAwal != null || JatuhTempo_NomorPKAkhir != null ||
		JatuhTempo_TanggalPKAwal != null || JatuhTempo_TanggalPKAkhir != null) {
		BoolJatuhTempo = true
		println("JatuhTempo akan dijalankan")
	}
	
	if (BoolPosisiRek) {
		println ("PosisiRek_RekeningAfiliasi : "+PosisiRek_RekeningAfiliasi)
		WebUI.check(findTestObject('Object Repository/Activity/ActivityMaintenanceRek_Object/check_PosisiRek_RekeningAfiliasi'))
		WebUI.setText(findTestObject('Object Repository/Activity/ActivityMaintenanceRek_Object/input_PosisiRek_Rekening Afiliasi_Menjadi'), PosisiRek_RekeningAfiliasi)
	}
	
	
	if (BoolRateBunga) {
		println ("PosisiRek_RateBunga : "+PosisiRek_RateBunga)
		WebUI.check(findTestObject('Object Repository/Activity/ActivityMaintenanceRek_Object/check_PosisiRek_RateBunga'))
		WebUI.setText(findTestObject('Object Repository/Activity/ActivityMaintenanceRek_Object/input_PosisiRek_RateBunga_Menjadi'), PosisiRek_RateBunga)
	}
	
	//println ("PosisiRek_RateBunga : "+PosisiRek_RateBunga)
	//WebUI.check(findTestObject('Object Repository/Activity/ActivityMaintenanceRek_Object/check_PosisiRek_RateBunga'))
	//WebUI.setText(findTestObject('Object Repository/Activity/ActivityMaintenanceRek_Object/input_PosisiRek_RateBunga_Menjadi'), PosisiRek_RateBunga)
	
	WebUI.setText(findTestObject('Object Repository/Activity/ActivityMaintenanceRek_Object/input_Nomor Rekening'), NoRek)
	
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
	
	if (BoolJatuhTempo) {
		WebUI.check(findTestObject('Object Repository/Activity/ActivityMaintenanceRek_Object/check_JatuhTempo'))
		println("JatuhTempo_NomorPKAwal :"+JatuhTempo_NomorPKAwal)
		WebUI.setText(findTestObject('Object Repository/Activity/ActivityMaintenanceRek_Object/input_Nomor PK Awal Menjadi'), JatuhTempo_NomorPKAwal != null ? JatuhTempo_NomorPKAwal : '')
		println("JatuhTempo_NomorPKAkhir :"+JatuhTempo_NomorPKAkhir)
		WebUI.setText(findTestObject('Object Repository/Activity/ActivityMaintenanceRek_Object/input_Nomor PK Akhir Menjadi'), JatuhTempo_NomorPKAkhir != null ? JatuhTempo_NomorPKAkhir : '')
		println("JatuhTempo_TanggalPKAwal :"+JatuhTempo_TanggalPKAwal)
		WebUI.setText(findTestObject('Object Repository/Activity/ActivityMaintenanceRek_Object/input_Tanggal PK Awal Menjadi'), JatuhTempo_TanggalPKAwal != null ? JatuhTempo_TanggalPKAwal : '')
		println("JatuhTempo_TanggalPKAkhir :"+JatuhTempo_TanggalPKAkhir)
		WebUI.setText(findTestObject('Object Repository/Activity/ActivityMaintenanceRek_Object/input_Tanggal PK Akhir Menjadi'), JatuhTempo_TanggalPKAkhir != null ? JatuhTempo_TanggalPKAkhir : '')
		println("JatuhTempo_PerpanjanganSementara :"+JatuhTempo_PerpanjanganSementara)
		WebUI.setText(findTestObject('Object Repository/Activity/ActivityMaintenanceRek_Object/input_JatuhTempo_Perpanjangan Sementara'), JatuhTempo_PerpanjanganSementara != null ? JatuhTempo_PerpanjanganSementara : '')
	}
	
	if (BoolMaksimum_Rekening) {
		println ("Maksimum_Rekening : "+Maksimum_Rekening)
		WebUI.check(findTestObject('Object Repository/Activity/ActivityMaintenanceRek_Object/check_MaksimumPinjaman'))
		WebUI.setText(findTestObject('Object Repository/Activity/ActivityMaintenanceRek_Object/input_Maksimum Rekening'), Maksimum_Rekening)
	}
	
	CustomKeywords.'custom.CustomKeywords.scrollToTop'()
	WebUI.click(findTestObject('Object Repository/Activity/ActivityMaintenanceRek_Object/input_Nomor Rekening'))
}

ActivityUtils.saveActivityAndCapture(NoTC, stepName, newDirectoryPath, numberCapture)