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
String excelFilePath = RunConfiguration.getProjectDir() + "/Data Files/NewSkenarioAllActivity.xlsx"
FileInputStream file = new FileInputStream(excelFilePath)
Workbook workbook = new XSSFWorkbook(file)
Sheet sheet3 = workbook.getSheet("RestrukturisasiRek")

// Cari berdasarkan TC
String NoRek = ""
String KodeRestruk = ""
String MetodeRestruk = ""
for (int i = 1; i <= sheet3.getLastRowNum(); i++) {
	Row row = sheet3.getRow(i)
	if (row != null && row.getCell(0).getStringCellValue() == NoTC) {
		NoRek = String.valueOf((long) row.getCell(5).getNumericCellValue())
		KodeRestruk = String.valueOf((long) row.getCell(6).getNumericCellValue())
		MetodeRestruk = String.valueOf((long) row.getCell(7).getNumericCellValue())
		break
	}
}
WebDriver driver = DriverFactory.getWebDriver()
WebUI.comment("TC: ${NoTC}")

WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Card Activity.png')
WebUI.click(findTestObject('Object Repository/COP/TabCard/a_Tab_PenutupanPendingRestrukturisasi'))
WebUI.click(findTestObject('Object Repository/COP/CardActivity/div_Card RestrukturisasiRek'))
WebUI.setText(findTestObject('Object Repository/Activity/ActivityRestrukturisasiRek_Object/input_NoRek'), NoRek)
String resultKodeResult = KodeRestrukList.find { it.contains(KodeRestruk) }
WebUI.selectOptionByLabel(findTestObject('Object Repository/Activity/ActivityRestrukturisasiRek_Object/select_Kode Restrukturisasi'), resultKodeResult, true)
String resultMetodeRestruk = MetodeRestrukList.find { it.contains(MetodeRestruk) }
WebUI.selectOptionByValue(findTestObject('Object Repository/Activity/ActivityRestrukturisasiRek_Object/select_Metode Restrukturisasi'), resultMetodeRestruk, true)

CustomKeywords.'custom.CustomKeywords.captureFullPageInSections'(newDirectoryPath+'/', numberCapture++ +'. form')

WebUI.click(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/button_Save'))
WebUI.waitForElementVisible(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/button_Save OK'), 30)
WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Sukses Save.png')
WebUI.click(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/button_Save OK'))
