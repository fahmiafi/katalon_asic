package custom

import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject

import com.kms.katalon.core.annotation.Keyword
import com.kms.katalon.core.checkpoint.Checkpoint
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling
import com.kms.katalon.core.testcase.TestCase
import com.kms.katalon.core.testdata.TestData
import com.kms.katalon.core.testobject.TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows

import internal.GlobalVariable
import custom.CustomKeywords
import logger.TestStepLogger

public class ActivityUtils {
	def static void saveActivityAndCapture(String NoTC, String stepName, String directoryPath, int numberCapture) {
		//		CustomKeywords.captureFullPageInSections(directoryPath + '/', numberCapture++ + '. Input Form')
		TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, numberCapture++, 'Input Form', directoryPath, true, true)
		WebUI.scrollToElement(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/button_Save'), 30)
		//		WebUI.takeScreenshot(directoryPath + '/' + numberCapture++ + '. Simpan.png')
		TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, numberCapture++, 'Simpan', directoryPath, true, false)
		WebUI.click(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/button_Save'))
		WebUI.waitForElementVisible(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/button_Save OK'), 30)
		//		WebUI.takeScreenshot(directoryPath + '/' + numberCapture++ + '. Berhasil disimpan.png')
		TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, numberCapture++, 'Berhasil disimpan', directoryPath, true, false)
		WebUI.click(findTestObject('Object Repository/Activity/ActivityBlokirRek_Object/button_Save OK'))
	}
}
