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

String NoTC = GlobalVariable.NoTC
String NoMemo = GlobalVariable.NoMemo
String Segmen = GlobalVariable.Segmen
String Pencairan = GlobalVariable.Pencairan
String UseCase = GlobalVariable.UseCase
String Skenario = GlobalVariable.Skenario
String Nominal = GlobalVariable.Nominal

String newDirectoryPath = GlobalVariable.newDirectoryPath
Integer numberCapture = 1


WebUI.waitForElementVisible(findTestObject('Object Repository/COP/a_Admin Kredit'), 30)

// Create Batch
WebUI.click(findTestObject('Object Repository/COP/a_Admin Kredit'))
WebUI.click(findTestObject('Object Repository/COP/a_Monitoring Batch Progress  Failed'))
WebUI.waitForElementVisible(findTestObject('Object Repository/COP/a_Create New Batch'), 30)
WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Create New Batch.png')
WebUI.click(findTestObject('Object Repository/COP/a_Create New Batch'))

String NamaDebitur = NoTC+' '+ UseCase + ' '+ Skenario +' '+ Pencairan
WebUI.setText(findTestObject('Object Repository/COP/input_NoTestKey'), NoMemo)
WebUI.setText(findTestObject('Object Repository/COP/input_NamaDebitur'), NamaDebitur.replaceAll(/\s+$/, '').replaceAll(/\s+/, ' '))
WebUI.selectOptionByValue(findTestObject('Object Repository/COP/DokUnderlying/select_Kategori Underlying'), '1', true)
WebUI.setText(findTestObject('Object Repository/COP/DokUnderlying/inputnama'), '123')

TestObject uploadField = findTestObject('Object Repository/COP/DokUnderlying/input_Dokumen Underlying_telexfile')
String filePath = 'D:\\BNI\\RPA\\.TESTING.pdf'
WebUI.uploadFile(uploadField, filePath)
WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Nama Debitur.png')
WebUI.click(findTestObject('Object Repository/COP/DokUnderlying/button_Upload'))
WebUI.delay(7)

// Add Activity
WebUI.scrollToElement(findTestObject('Object Repository/COP/DokUnderlying/button_Upload'), 30)
WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Add Activity.png')
WebUI.click(findTestObject('Object Repository/COP/button_Add Activity'))

// Choose Card
if (UseCase == "Pemindahbukuan") {
	WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Card Activity.png')
	WebUI.click(findTestObject('Object Repository/COP/CardActivity/div_card_pinbuk'))
} else if (UseCase == "Asuransi") {
	WebUI.click(findTestObject('Object Repository/COP/TabCard/a_Tab_AsuransiPembukaan'))
	WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Card Activity.png')
	WebUI.click(findTestObject('Object Repository/COP/CardActivity/div_Card Asuransi Jaminan'))
} else if (UseCase == "Bucket Adjusment") {
	WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Card Activity.png')
	WebUI.click(findTestObject('Object Repository/COP/CardActivity/div_Card Bucket Adjustment'))
}
