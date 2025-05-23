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
String Pencairan = GlobalVariable.Pencairan
String Skenario = GlobalVariable.Skenario
String newDirectoryPath = GlobalVariable.newDirectoryPath
Integer numberCapture = 1

// Path ke file Excel
String excelFilePath = RunConfiguration.getProjectDir() + "/Data Files/NewSkenarioAllActivity.xlsx"
FileInputStream file = new FileInputStream(excelFilePath)
Workbook workbook = new XSSFWorkbook(file)
Sheet sheet3 = workbook.getSheet("Pemindahbukuan")

// Cari berdasarkan TC
String RekDebit = ""
String RekKredit = ""
String Nominal = ""
for (int i = 1; i <= sheet3.getLastRowNum(); i++) {
	Row row = sheet3.getRow(i)
	if (row != null && row.getCell(0).getStringCellValue() == NoTC) {
//		RekDebit = String.valueOf((long) row.getCell(5).getNumericCellValue())
//		RekKredit = String.valueOf((long) row.getCell(6).getNumericCellValue())
	
		RekDebit = ExcelHelper.getCellValueAsString(row, 5)
		RekKredit = ExcelHelper.getCellValueAsString(row, 6)
		Nominal = String.valueOf((long) row.getCell(7).getNumericCellValue())
		break
	}
}

WebUI.comment("TC: ${NoTC}; Pencairan: ${Pencairan} ; RekDebit: ${RekDebit}; RekKredit ${RekKredit} Nominal: ${Nominal}")

WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Card Activity.png')
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
WebUI.setText(findTestObject('Object Repository/Activity/ActivtyPemindahbukuan/textarea_DetailPemindahan.Narasi'), "ASIC-"+DataTimeNow+"-"+NoTC)
WebUI.setText(findTestObject('Object Repository/Activity/ActivtyPemindahbukuan/textarea_DetailPemindahan.NarasiTambahan'), 'Test Narasi 2')

CustomKeywords.'custom.CustomKeywords.captureFullPageInSections'(newDirectoryPath+'/', numberCapture++ +'. form')

WebUI.click(findTestObject('Object Repository/COP/button_Save'))
WebUI.waitForElementVisible(findTestObject('Object Repository/COP/button_OK (1)'), 30)
WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Sukses Save.png')
WebUI.click(findTestObject('Object Repository/COP/button_OK (1)'))
