package stepCapture

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
import com.kms.katalon.core.webui.driver.DriverFactory
import org.openqa.selenium.JavascriptExecutor
import org.openqa.selenium.interactions.Actions
import org.openqa.selenium.Keys
import com.kms.katalon.core.testobject.ConditionType

import internal.GlobalVariable
import logger.TestStepLogger

public class StepCaptureHelper {
	@Keyword
	def scrollCapture(String stepName, String NoTC, int numberCapture, String imgName) {
		String newDirectoryPath = GlobalVariable.PathCapture+"\\"+NoTC+"\\"+stepName
		int page = 1
		def driver = DriverFactory.getWebDriver()
		def scrollStep = 650
		def scrollInterval = 100

		def scrollStepByStep = {
			((JavascriptExecutor) driver).executeScript("document.getElementById('viewerContainer').scrollBy(0, arguments[0]);", scrollStep)

			def scrollTop = ((JavascriptExecutor) driver).executeScript("return document.getElementById('viewerContainer').scrollTop;")
			def clientHeight = ((JavascriptExecutor) driver).executeScript("return document.getElementById('viewerContainer').clientHeight;")
			def scrollHeight = ((JavascriptExecutor) driver).executeScript("return document.getElementById('viewerContainer').scrollHeight;")

			if (scrollTop + clientHeight >= scrollHeight - 2) {
				println("Sudah mencapai bagian bawah halaman.")
				return true
			}
			return false
		}

		List<String> imageFiles = []
		String filename = numberCapture + '. '+ imgName +'_' + page++ + '.png'
		imageFiles << stepName+"/"+filename
		WebUI.takeScreenshot(newDirectoryPath + '/' + filename)
		while (true) {
			def isBottomReached = scrollStepByStep()
			filename = numberCapture + '. '+ imgName +'_' + page++ + '.png'
			WebUI.takeScreenshot(newDirectoryPath + '/' + filename)
			if (isBottomReached) {
				break
			}
			imageFiles << stepName+"/"+filename
			WebUI.delay(scrollInterval / 1000)
		}
		String[] imageFilesArray = imageFiles.toArray(new String[0])
		TestStepLogger.addOutputWithUserAndWithOutCapture(NoTC, 'Maker', 4, imgName, imageFilesArray)
	}

	@Keyword
	def pageCapture(String stepName, String NoTC, int numberCapture, String imgName) {
		String newDirectoryPath = GlobalVariable.PathCapture+"\\"+NoTC+"\\"+stepName
		def driver = DriverFactory.getWebDriver()
		def actions = new Actions(driver)

		String numPages = WebUI.executeJavaScript("return document.querySelector('#numPages').innerText;", null)
		def inputElement = new TestObject('dynamicInput')
		inputElement.addProperty('xpath', ConditionType.EQUALS, "//input[@id='pageNumber' and @type='number']")

		WebUI.comment("Teks yang diambil: " + numPages)

		String pageNumberStr = numPages.replaceAll("[^0-9]", "")
		int pageNumber = Integer.parseInt(pageNumberStr)

		List<String> imageFiles = []
		String filename = ""
		for (int j = 1; j <= pageNumber; j++) {
			int page = j
			filename = numberCapture++ + '. '+ imgName +'_' + page + '.png'

			WebUI.click(inputElement)

			actions.keyDown(Keys.CONTROL).sendKeys("a").keyUp(Keys.CONTROL).sendKeys(Keys.DELETE).perform()
			WebUI.delay(1)

			WebUI.setText(inputElement, page.toString())

			actions.sendKeys(Keys.ENTER).perform()
			WebUI.delay(2)

			WebUI.takeScreenshot(newDirectoryPath + '/' + filename)
			imageFiles << stepName+"/"+filename
		}
		String[] imageFilesArray = imageFiles.toArray(new String[0])
		TestStepLogger.addOutputWithUserAndWithOutCapture(NoTC, 'Maker', 4, imgName, imageFilesArray)
	}
}
