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

import org.openqa.selenium.WebDriver
import com.kms.katalon.core.webui.driver.DriverFactory
import com.kms.katalon.core.configuration.RunConfiguration
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileInputStream
import com.kms.katalon.core.testobject.ConditionType
import org.openqa.selenium.By
import org.openqa.selenium.WebElement

import utils.LogHelper
import excel.ExcelHelper

String stepName = 'Maker'

// Path ke file Excel
String excelFilePath = RunConfiguration.getProjectDir() + "/Data Files/SkenarioValidasiPK.xlsx"
FileInputStream file = new FileInputStream(excelFilePath)
Workbook workbook = new XSSFWorkbook(file)
Sheet sheetSkenario = workbook.getSheet("Skenario Pembukaan Excel")

WebUI.openBrowser('')
WebUI.navigateToUrl('http://192.168.174.45/Login')
WebUI.maximizeWindow()

// Loop melalui data di Sheet1
for (int i = 1; i <= sheetSkenario.getLastRowNum(); i++) {
	Row row = sheetSkenario.getRow(i)
	String checkRunning = row.getCell(1).getStringCellValue()
	if (row != null && checkRunning == "Y") {
		String NoTC = ExcelHelper.getCellValueAsString(row, 0)
		String IsRunning = ExcelHelper.getCellValueAsString(row, 1)
		String UseCase = ExcelHelper.getCellValueAsString(row, 2)
		String Metode = ExcelHelper.getCellValueAsString(row, 3)
		String Skenario = ExcelHelper.getCellValueAsString(row, 4)
		String TestStep = ExcelHelper.getCellValueAsString(row, 5)
		String ExpectedResult = ExcelHelper.getCellValueAsString(row, 6)
		String DataTest = ExcelHelper.getCellValueAsString(row, 7)
		String NomorPKAwal = ExcelHelper.getCellValueAsString(row, 8)
		String NomorPKAkhir = ExcelHelper.getCellValueAsString(row, 9)
		String TanggalPKAwal = ExcelHelper.getCellValueAsString(row, 10)
		String TanggalPKAkhir = ExcelHelper.getCellValueAsString(row, 11)
		String MakerNpp = ExcelHelper.getCellValueAsString(row, 12)
		String MakerPassword = ExcelHelper.getCellValueAsString(row, 13)
		String FileExcel = ExcelHelper.getCellValueAsString(row, 18)
		String Cif = DataTest
		
		String Nominal = "1000000000000"
		
		String newDirectoryPath = GlobalVariable.PathCapture+"\\"+NoTC+"\\"+stepName
		GlobalVariable.newDirectoryPath = newDirectoryPath
		Integer numberCapture = 1
		
		// Login
		WebUI.setText(findTestObject('Object Repository/Login/inputtxtUsername'), MakerNpp)
		WebUI.setText(findTestObject('Object Repository/Login/inputtxtPassword'), MakerPassword)
		WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Login sebagai maker.png')
		WebUI.click(findTestObject('Object Repository/Login/button_Sign In'))
		
		WebUI.waitForElementVisible(findTestObject('Object Repository/COP/a_Admin Kredit'), 30)
		
		// Create Batch
		WebUI.click(findTestObject('Object Repository/COP/a_Admin Kredit'))
		WebUI.click(findTestObject('Object Repository/COP/a_Monitoring Batch Progress  Failed'))
		// Buat TestObject dinamis untuk elemen loading
		TestObject loadingPanel = new TestObject().tap {
			addProperty("xpath", ConditionType.EQUALS, "//div[contains(@class, 'jsgrid-load-panel')]")
		}
		
		// Tunggu maksimal 30 detik hingga loading tidak terlihat
		WebUI.waitForElementNotVisible(loadingPanel, 30)
		WebUI.waitForElementVisible(findTestObject('Object Repository/COP/a_Create New Batch'), 30)
		WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Create New Batch.png')
		WebUI.click(findTestObject('Object Repository/COP/a_Create New Batch'))
		
		String NamaDebitur = NoTC+' '+ Skenario
		// Potong jika panjang lebih dari 100 karakter
		if (NamaDebitur.length() > 99) {
			NamaDebitur = NamaDebitur.substring(0, 99)
		}
		WebUI.waitForElementVisible(findTestObject('Object Repository/COP/DokUnderlying/select_Kategori Underlying'), 30)
		WebUI.setText(findTestObject('Object Repository/COP/input_NoTestKey'), NoTC+"-01")
		WebUI.setText(findTestObject('Object Repository/COP/input_NamaDebitur'), NamaDebitur.replaceAll(/\s+$/, '').replaceAll(/\s+/, ' '))
		WebUI.delay(2)
		WebUI.selectOptionByValue(findTestObject('Object Repository/COP/DokUnderlying/select_Kategori Underlying'), '1', true)
		WebUI.setText(findTestObject('Object Repository/COP/DokUnderlying/inputnama'), '123')
		
		TestObject uploadField = findTestObject('Object Repository/COP/DokUnderlying/input_Dokumen Underlying_telexfile')
		String filePath = 'D:\\BNI\\RPA\\.TESTING.pdf'
		WebUI.uploadFile(uploadField, filePath)
		WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Input No. Memo, Nama Debitur dan Upload Dokumen Underlying.png')
		WebUI.click(findTestObject('Object Repository/COP/DokUnderlying/button_Upload'))
//		WebUI.delay(7)
		WebUI.waitForElementVisible(findTestObject('Object Repository/COP/DokUnderlying/td_No_Memo_Underlying'), 30)
		
		// Add Activity
		WebUI.scrollToElement(findTestObject('Object Repository/COP/DokUnderlying/label_Flag Batch'), 30)
		WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Add Activity.png')
		WebUI.click(findTestObject('Object Repository/COP/button_Add Activity'))
		
		int NumberAct = 1
		if (UseCase == 'Pembukaan Rekening') {
			newDirectoryPath = GlobalVariable.PathCapture+"\\"+NoTC+"\\"+stepName+"\\Form Activity-"+NumberAct
			GlobalVariable.newDirectoryPath = newDirectoryPath
			
			WebUI.click(findTestObject('Object Repository/COP/TabCard/a_Tab_AsuransiPembukaan'))
			WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Pilih Use Case '+ UseCase +'.png')
			WebUI.click(findTestObject('Object Repository/COP/CardActivity/div_Card_Pembukaan Rekening Pinjaman'))
			
			if (Metode == 'Input Excel') {
				WebUI.click(findTestObject('Object Repository/COP/input_Excel'))
				TestObject uploadExel = findTestObject('Object Repository/COP/label_Upload Excel Activity')
				String filePathExcel = 'D:\\BNI\\RPA\\CR Antasena, Enhance Robot\\Data Excel\\'+FileExcel
				WebUI.uploadFile(uploadExel, filePathExcel)
			}
			else {				
				WebUI.setText(findTestObject('Object Repository/Activity/ActivityPembukaanRek_Object/input_CIF'), Cif)
				WebUI.click(findTestObject('Object Repository/Activity/ActivityPembukaanRek_Object/a_Search Cif'))
				//WebUI.delay(3)
				//WebUI.click(findTestObject('Object Repository/Activity/ActivityPembukaanRek_Object/button_OK_Alert Form Pembukaan'))
				TestObject spinner = new TestObject()
				spinner.addProperty("xpath", ConditionType.EQUALS, "//div[contains(@class, 'loadingSpinner')]")
				WebUI.waitForElementNotVisible(spinner, 30)
				WebUI.delay(2)
				
				WebDriver driver = DriverFactory.getWebDriver()
				WebUI.click(findTestObject('Object Repository/Activity/ActivityPembukaanRek_Object/span_Cabang Pembuka'))
				WebUI.delay(1)
				List<WebElement> subCabPembukaOptions = driver.findElements(By.xpath("//li[contains(@class,'select2-results__option') and contains(text(),'760 : DIVISI OPERASIONAL - JPC')]"))
				subCabPembukaOptions[0].click()
				WebUI.click(findTestObject('Object Repository/Activity/ActivityPembukaanRek_Object/span_Sub Kategori'))
				WebUI.delay(1)
				List<WebElement> subKategoriOptions = driver.findElements(By.xpath("//li[contains(@class,'select2-results__option') and contains(text(),'Efektif IDR')]"))
				subKategoriOptions[0].click()
				
				println("NomorPKAwal :"+NomorPKAwal)
				WebUI.setText(findTestObject('Object Repository/Activity/ActivityPembukaanRek_Object/input_NoPK_Awal'), NomorPKAwal != null ? NomorPKAwal : '')
				println("NomorPKAkhir :"+NomorPKAkhir)
				WebUI.setText(findTestObject('Object Repository/Activity/ActivityPembukaanRek_Object/input_NoPK_Akhir'), NomorPKAkhir != null ? NomorPKAkhir : '')
				println("TanggalPKAwal :"+TanggalPKAwal)
				WebUI.setText(findTestObject('Object Repository/Activity/ActivityPembukaanRek_Object/input_TanggalPK_Awal'), TanggalPKAwal != null ? TanggalPKAwal : '')
				println("TanggalPKAkhir :"+TanggalPKAkhir)
				WebUI.setText(findTestObject('Object Repository/Activity/ActivityPembukaanRek_Object/input_TanggalPK_Akhir'), TanggalPKAkhir != null ? TanggalPKAkhir : '')
				
				WebUI.setText(findTestObject('Object Repository/Activity/ActivityPembukaanRek_Object/input_Maksimum Kredit'), Nominal.replaceAll("[^0-9]", ""))
				
				WebUI.click(findTestObject('Object Repository/Activity/ActivityPembukaanRek_Object/input_Tidak ada Pencairan'))
			}
//			WebUI.delay(20)
			CustomKeywords.'custom.CustomKeywords.captureFullPageInSections'(newDirectoryPath+'/', numberCapture++ +'. Input Form')
			
			WebUI.scrollToElement(findTestObject('Object Repository/Activity/ActivityPembukaanRek_Object/button_Save_Form_Pembukaan'), 30)
			WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Simpan.png')
			WebUI.click(findTestObject('Object Repository/Activity/ActivityPembukaanRek_Object/button_Save_Form_Pembukaan'))
			
			TestObject AlertTitle = findTestObject('Object Repository/ValidasiPK/AlertTitle')
			TestObject AlertMessage = findTestObject('Object Repository/ValidasiPK/AlertMessage')
			TestObject AlertConfirm = findTestObject('Object Repository/ValidasiPK/AlertConfirm')
			// Tunggu sampai alert muncul
			WebUI.waitForElementVisible(AlertTitle, 10)
			
			// Ambil teks pesan alert
			String Title = WebUI.getText(AlertTitle)
			String Message = WebUI.getText(AlertMessage)
			println("Title : " + Title)
			println("Message :" + Message)
			WebUI.delay(3)
			WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. PROCEDURE - Muncul popup message.png')
			if (Title == 'Success') {
				WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Berhasil disimpan.png')
				WebUI.click(AlertConfirm)
				
				
				newDirectoryPath = GlobalVariable.PathCapture+"\\"+NoTC+"\\"+stepName
				GlobalVariable.newDirectoryPath = newDirectoryPath
				
				WebUI.scrollToElement(findTestObject('Object Repository/COP/DokUnderlying/label_Flag Batch'), 30)
				// inquiry
				WebUI.waitForElementVisible(findTestObject('Object Repository/COP/button_inquiry'), 30)
				WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Activity berhasil ditambahkan, status New.png')
				WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Inquiry pada table List Aktivitas kolom Action atau Inquiry All.png')
				WebUI.click(findTestObject('Object Repository/COP/button_Inquiry All'))
				WebUI.waitForElementVisible(findTestObject('Object Repository/COP/button_OK_inquiry'), 30)
				WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Inquiry berhasil.png')
				WebUI.click(findTestObject('Object Repository/COP/button_OK_inquiry'))
				WebUI.delay(1)
				WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Status Inquiring.png')
			}
			else {
				WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. OUTPUT - Muncul popup message Error.png')
				WebUI.click(AlertConfirm)
			}
		}
		
		
		
		// Logout
		WebUI.click(findTestObject('Object Repository/Login/i_User Logout'))
		WebUI.click(findTestObject('Object Repository/Login/a_Logout'))
		WebUI.delay(2)
	}
}

// Tutup
workbook.close()
file.close()
WebUI.closeBrowser()