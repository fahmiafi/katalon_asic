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

List<String> KodeRestrukList = [
	'0: Remove',
	'1: Get commitment',
	'2: Stand still status',
	'3. Appoint ext financial/legal cons.',
	'4: Due diligence',
	'5: Negotiate restruct',
	'6: Finalise restruct',
	'7: Implementation & monitoring',
	'8: Gagal restruk',
]
List<String> MetodeRestrukList = [
	'1: Penurunan suku bunga kredit',
	'2: Perpanjangan jangka waktu kredit',
	'3: Pengurangan tunggakan pokok kredit',
	'4: Pengurangan tunggakan bunga kredit',
	'5: Penambahan fasilitas kredit',
	'6: Konversi kredit menjadi Penyertaan Modal Sementara',
	'7: Penambahan fasilitas kredit dan pengurangan tunggakan bunga kredit',
	'8: Penambahan fasilitas kredit dan perpanjangan jk waktu kredit',
	'9: Penambahan fasilitas kredit dan penurunan suku bunga',
	'0: Penambahan fasilitas kredit, pengurangan tunggakan bunga kredit dan penurunan suku bunga kredit',
	'A: Penambahan fasilitas kredit,  pengurangan tunggakan bunga kredit, dan perpanjangan jangka waktu kredit',
	'B: Lainnya',
]
// Path ke file Excel
String excelFilePath = RunConfiguration.getProjectDir() + GlobalVariable.PathDataExcel
FileInputStream file = new FileInputStream(excelFilePath)
Workbook workbook = new XSSFWorkbook(file)
Sheet sheetAct = workbook.getSheet("Act Restrukturisasi Rek")

// Cari berdasarkan TC
String TanggalRestruk          = ""
String Frekuensi               = ""
String KodeRM                  = ""
String Deskripsi               = ""
String NoRek                   = ""
String KodeRestruk             = ""
String MetodeRestruk           = ""
String KolektabilitasRekening  = ""
String FlagStimulus            = ""
for (int i = 2; i <= sheetAct.getLastRowNum(); i++) {
	Row row = sheetAct.getRow(i)
	if (row != null && ExcelHelper.getCellValueAsString(row, 0) == NoTC && ExcelHelper.getCellValueAsString(row, 1) == Seq) {
		TanggalRestruk          = ExcelHelper.getCellValueAsString(row, 4)
		Frekuensi               = ExcelHelper.getCellValueAsString(row, 5)
		KodeRM                  = ExcelHelper.getCellValueAsString(row, 6)
		Deskripsi               = ExcelHelper.getCellValueAsString(row, 7)
		NoRek                   = ExcelHelper.getCellValueAsString(row, 8)
		KodeRestruk             = ExcelHelper.getCellValueAsString(row, 9)
		MetodeRestruk           = ExcelHelper.getCellValueAsString(row, 10)
		KolektabilitasRekening  = ExcelHelper.getCellValueAsString(row, 11)
		FlagStimulus            = ExcelHelper.getCellValueAsString(row, 12)
		break
	}
}
WebDriver driver = DriverFactory.getWebDriver()
WebUI.comment("TC: ${NoTC}")

WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Card Activity.png')
WebUI.click(findTestObject('Object Repository/COP/TabCard/a_Tab_PenutupanPendingRestrukturisasi'))
WebUI.click(findTestObject('Object Repository/COP/CardActivity/div_Card RestrukturisasiRek'))
WebUI.delay(3)
WebUI.setText(findTestObject('Object Repository/Activity/ActivityRestrukturisasiRek_Object/input_NoRek'), NoRek)
WebUI.setText(findTestObject('Object Repository/Activity/ActivityRestrukturisasiRek_Object/input_TglRestruk'), TanggalRestruk)
//WebUI.click(findTestObject('Object Repository/Activity/ActivityRestrukturisasiRek_Object/input_Frekuensi'))
WebUI.setText(findTestObject('Object Repository/Activity/ActivityRestrukturisasiRek_Object/input_Frekuensi'), Frekuensi)
WebUI.setText(findTestObject('Object Repository/Activity/ActivityRestrukturisasiRek_Object/input_KodeRM'), KodeRM)
WebUI.setText(findTestObject('Object Repository/Activity/ActivityRestrukturisasiRek_Object/textarea_Deskripsi'), Deskripsi)
String resultKodeResult = KodeRestrukList.find { it.contains(KodeRestruk) }
WebUI.selectOptionByLabel(findTestObject('Object Repository/Activity/ActivityRestrukturisasiRek_Object/select_Kode Restrukturisasi'), resultKodeResult, true)
WebUI.delay(3)
String resultMetodeRestruk = MetodeRestrukList.find { it.contains(MetodeRestruk) }
WebUI.selectOptionByValue(findTestObject('Object Repository/Activity/ActivityRestrukturisasiRek_Object/select_Metode Restrukturisasi'), resultMetodeRestruk, true)
WebUI.selectOptionByLabel(findTestObject('Object Repository/Activity/ActivityRestrukturisasiRek_Object/select_Kolektabilitas Rekening'), KolektabilitasRekening, true)
WebUI.selectOptionByLabel(findTestObject('Object Repository/Activity/ActivityRestrukturisasiRek_Object/select_FlagStimulus'), FlagStimulus, true)

CustomKeywords.'custom.CustomKeywords.scrollToTop'()
WebUI.click(findTestObject('Object Repository/Activity/ActivityRestrukturisasiRek_Object/input_NoRek'))

CustomKeywords.'custom.CustomKeywords.captureFullPageInSections'(newDirectoryPath+'/', numberCapture++ +'. Input Form')

WebUI.scrollToElement(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/button_Save'), 30)
WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Simpan.png')
WebUI.click(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/button_Save'))
WebUI.waitForElementVisible(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/button_Save OK'), 30)
WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Berhasil disimpan.png')
WebUI.click(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/button_Save OK'))

