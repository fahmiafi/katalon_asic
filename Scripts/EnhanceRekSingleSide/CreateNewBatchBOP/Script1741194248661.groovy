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


WebUI.waitForElementVisible(findTestObject('Object Repository/BOP/CreateNewBatch/p_Admin Kredit_Menu'), 30)

// Create Batch
WebUI.click(findTestObject('Object Repository/BOP/CreateNewBatch/p_Admin Kredit_Menu'))
WebUI.click(findTestObject('Object Repository/BOP/CreateNewBatch/a_Admin Kredit_SubMenu'))
WebUI.waitForElementVisible(findTestObject('Object Repository/BOP/CreateNewBatch/i_Admin Kredit_Create New Batch'), 30)
WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Create New Batch.png')
WebUI.click(findTestObject('Object Repository/BOP/CreateNewBatch/i_Admin Kredit_Create New Batch'))

String NamaDebitur = NoTC+' '+ UseCase + ' '+ Skenario +' '+ Pencairan
WebUI.setText(findTestObject('Object Repository/BOP/CreateNewBatch/input_NoTestKey'), NoMemo)
WebUI.setText(findTestObject('Object Repository/BOP/CreateNewBatch/input_NamaDebitur'), NamaDebitur.replaceAll(/\s+$/, '').replaceAll(/\s+/, ' '))

WebUI.scrollToElement(findTestObject('Object Repository/BOP/CreateNewBatch/i_AddActivity'), 30)
TestObject uploadField = findTestObject('Object Repository/BOP/CreateNewBatch/input_UploadTelex')
String filePath = 'D:\\BNI\\RPA\\.TESTING.pdf'
WebUI.uploadFile(uploadField, filePath)
WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Nama Debitur.png')
WebUI.delay(7)

// Add Activity
WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Add Activity.png')
WebUI.click(findTestObject('Object Repository/BOP/CreateNewBatch/i_AddActivity'))

// Choose Card
WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Card Activity.png')
if (UseCase == "Pemindahbukuan") {
	WebUI.click(findTestObject('Object Repository/BOP/CardActivity/h5_Pemindahbukuan Dana'))
} else if (UseCase == "Asuransi") {
	WebUI.click(findTestObject('Object Repository/BOP/CardActivity/h5_Asuransi Jaminan'))
} else if (UseCase == "Bucket Adjusment") {
	WebUI.click(findTestObject('Object Repository/BOP/CardActivity/h5_Bucket Adjustment'))
}