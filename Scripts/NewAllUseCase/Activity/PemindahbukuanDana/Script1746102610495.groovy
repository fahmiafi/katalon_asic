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
import java.text.SimpleDateFormat
import java.util.Date
import excel.ExcelHelper

String testCaseName = RunConfiguration.getExecutionSourceName()
String NoTC = GlobalVariable.NoTC
String Seq = GlobalVariable.Seq
String Pencairan = GlobalVariable.Pencairan
String newDirectoryPath = GlobalVariable.newDirectoryPath
Integer numberCapture = 1
String UseCase = GlobalVariable.UseCase

// Path ke file Excel
String excelFilePath = RunConfiguration.getProjectDir() + GlobalVariable.PathDataExcel
FileInputStream file = new FileInputStream(excelFilePath)
Workbook workbook = new XSSFWorkbook(file)
Sheet sheetAct = workbook.getSheet("Act Pemindahbukuan")

// Cari berdasarkan TC
String RekDebit = ""
String RekKredit = ""
String Nominal = ""
for (int i = 2; i <= sheetAct.getLastRowNum(); i++) {
	Row row = sheetAct.getRow(i)
	if (row != null && ExcelHelper.getCellValueAsString(row, 0) == NoTC && ExcelHelper.getCellValueAsString(row, 1) == Seq) {
		RekDebit 	= ExcelHelper.getCellValueAsString(row, 4)
		RekKredit 	= ExcelHelper.getCellValueAsString(row, 5)
		Nominal 	= ExcelHelper.getCellValueAsString(row, 6)
		break
	}
}

WebUI.comment("TC: ${NoTC}; Pencairan: ${Pencairan} ; RekDebit: ${RekDebit}; RekKredit ${RekKredit} Nominal: ${Nominal}")

WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Pilih Use Case '+ UseCase +'.png')
WebUI.click(findTestObject('Object Repository/COP/CardActivity/div_card_pinbuk'))

if (Pencairan == 'Pencairan Pertama') {
	WebUI.click(findTestObject('Object Repository/Activity/ActivtyPemindahbukuan/label_Pencairan Pertama'))
}
else if (Pencairan == 'Pencairan ke 2 dan seterusnya') {
	WebUI.click(findTestObject('Object Repository/Activity/ActivtyPemindahbukuan/label_Pencairan Kedua dan Sterusnya'))
}
else {
	WebUI.click(findTestObject('Object Repository/Activity/ActivtyPemindahbukuan/label_Non Pencairan'))
}


WebUI.setText(findTestObject('Object Repository/Activity/ActivtyPemindahbukuan/input_NoRekDebit'), RekDebit)
WebUI.setText(findTestObject('Object Repository/Activity/ActivtyPemindahbukuan/input_NoRekKredit'), RekKredit)
WebUI.setText(findTestObject('Object Repository/Activity/ActivtyPemindahbukuan/input_DetailPemindahan.Nominal'), Nominal.replaceAll("[^0-9]", ""))
// Ambil tanggal dan waktu sekarang
Date now = new Date()
SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss")
String DataTimeNow = sdf.format(now)
WebUI.setText(findTestObject('Object Repository/Activity/ActivtyPemindahbukuan/textarea_DetailPemindahan.Narasi'), "ASIC-"+DataTimeNow+"-"+NoTC+"-"+Seq)
WebUI.setText(findTestObject('Object Repository/Activity/ActivtyPemindahbukuan/textarea_DetailPemindahan.NarasiTambahan'), 'Test Narasi 2')

CustomKeywords.'custom.CustomKeywords.captureFullPageInSections'(newDirectoryPath+'/', numberCapture++ +'. Input Form')

WebUI.scrollToElement(findTestObject('Object Repository/COP/button_Save'), 30)
WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Simpan.png')
WebUI.click(findTestObject('Object Repository/COP/button_Save'))
WebUI.waitForElementVisible(findTestObject('Object Repository/COP/button_OK (1)'), 30)
WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Berhasil disimpan.png')
WebUI.click(findTestObject('Object Repository/COP/button_OK (1)'))
