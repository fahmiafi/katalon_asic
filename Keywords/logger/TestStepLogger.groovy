package logger

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

import com.kms.katalon.core.webui.driver.DriverFactory
import com.kms.katalon.core.util.KeywordUtil
import org.openqa.selenium.*
import javax.imageio.ImageIO
import java.awt.image.BufferedImage
import groovy.json.*
import java.nio.file.*

class TestStepLogger {
	static String logFilePath = "D:/katalon/ss/log.json"

	/**
	 * Tambahkan step dengan screenshot dan narasi berdasarkan user, mendukung capture full page
	 */
	static void addStepWithUserAndCapture(String noTC, String user, int NumberCapture, int narationLevel, String naration, String saveDir, boolean doCapture, boolean fullPage) {
		String logDirPath = GlobalVariable.PathCapture + "/" + noTC
		String captureDirPath = GlobalVariable.PathCapture + "/" + noTC + "/" + saveDir
		File logDir = new File(logDirPath)
		if (!logDir.exists()) {
			logDir.mkdirs()
		}
		File file = new File(logDir, "log.json")
		if (!file.exists()) {
			file.text = '[]'
		}

		def parser = new JsonSlurper()
		def logData = parser.parseText(file.text)

		def testCase = logData.find { it.NoTC == noTC }
		if (!testCase) {
			testCase = [NoTC: noTC, procedures: [], outputs: []]
			logData << testCase
		}

		def procedure = testCase.procedures.find { it.user == user }
		if (!procedure) {
			procedure = [user: user, steps: []]
			testCase.procedures << procedure
		}

		def safeNaration = naration.replaceAll(/[^a-zA-Z0-9 _-]/, '').replaceAll(/\s+/, '-')
		File directory = new File(captureDirPath)
		if (!directory.exists()) {
			directory.mkdirs()
		}

		List<String> imageFiles = []
		if (doCapture) {
			if (fullPage) {
				imageFiles = captureFullPageImagesForStep(noTC, saveDir, NumberCapture, safeNaration)
			} else {
				String filename = "${NumberCapture}. ${safeNaration}.png"
				WebUI.takeScreenshot("${captureDirPath}/${filename}")
				imageFiles << saveDir + "/" + filename
			}
		}

		procedure.steps << [
			datetime : new Date().format("yyyy-MM-dd HH:mm:ss"),
			narationLevel : narationLevel,
			naration: naration,
			images  : imageFiles
		]

		def jsonOutput = JsonOutput.prettyPrint(JsonOutput.toJson(logData))
		file.text = jsonOutput
	}
	
	static void addOutputWithUserAndCapture(String noTC, String user, int NumberCapture, int narationLevel, String naration, String saveDir, boolean doCapture, boolean fullPage) {
		String logDirPath = GlobalVariable.PathCapture + "/" + noTC
		String captureDirPath = GlobalVariable.PathCapture + "/" + noTC + "/" + saveDir
		File logDir = new File(logDirPath)
		if (!logDir.exists()) {
			logDir.mkdirs()
		}
		File file = new File(logDir, "log.json")
		if (!file.exists()) {
			file.text = '[]'
		}

		def parser = new JsonSlurper()
		def logData = parser.parseText(file.text)

		def testCase = logData.find { it.NoTC == noTC }
		if (!testCase) {
			testCase = [NoTC: noTC, procedures: [], outputs: []]
			logData << testCase
		}

		def output = testCase.outputs

		def safeNaration = naration.replaceAll(/[^a-zA-Z0-9 _-]/, '').replaceAll(/\s+/, '-')
		File directory = new File(captureDirPath)
		if (!directory.exists()) {
			directory.mkdirs()
		}

		List<String> imageFiles = []
		if (doCapture) {
			if (fullPage) {
				imageFiles = captureFullPageImagesForStep(noTC, saveDir, NumberCapture, safeNaration)
			} else {
				String filename = "${NumberCapture}. ${safeNaration}.png"
				WebUI.takeScreenshot("${captureDirPath}/${filename}")
				imageFiles << saveDir + "/" + filename
			}
		}

		output << [
			datetime : new Date().format("yyyy-MM-dd HH:mm:ss"),
			narationLevel : narationLevel,
			naration: naration,
			images  : imageFiles
		]

		def jsonOutput = JsonOutput.prettyPrint(JsonOutput.toJson(logData))
		file.text = jsonOutput
	}

	static void addStepWithUserAndWithOutCapture(String noTC, String user, int narationLevel, String naration, String[] imageFiles) {
		String logDirPath = GlobalVariable.PathCapture + "/" + noTC
		File logDir = new File(logDirPath)
		if (!logDir.exists()) {
			logDir.mkdirs()
		}
		File file = new File(logDir, "log.json")
		if (!file.exists()) {
			file.text = '[]'
		}

		def parser = new JsonSlurper()
		def logData = parser.parseText(file.text)

		def testCase = logData.find { it.NoTC == noTC }
		if (!testCase) {
			testCase = [NoTC: noTC, procedures: []]
			logData << testCase
		}

		def procedure = testCase.procedures.find { it.user == user }
		if (!procedure) {
			procedure = [user: user, steps: []]
			testCase.procedures << procedure
		}

		procedure.steps << [
			datetime : new Date().format("yyyy-MM-dd HH:mm:ss"),
			narationLevel : narationLevel,
			naration: naration,
			images  : imageFiles
		]

		def jsonOutput = JsonOutput.prettyPrint(JsonOutput.toJson(logData))
		file.text = jsonOutput
	}

	/**
	 * Fungsi untuk capture satu halaman penuh, dikembalikan dalam bentuk list nama file
	 */
	static List<String> captureFullPageImagesForStep(String noTC, String outputDirectory, int NumberCapture, String baseFilename) {
		String captureDirPath = GlobalVariable.PathCapture + "/" + noTC + "/" + outputDirectory

		WebDriver driver = DriverFactory.getWebDriver()
		JavascriptExecutor jsExecutor = (JavascriptExecutor) driver

		Long totalHeight = (Long) jsExecutor.executeScript("return document.body.scrollHeight")
		Long viewportHeight = (Long) jsExecutor.executeScript("return window.innerHeight")

		int steps = (int) Math.ceil(totalHeight / (double) viewportHeight)
		int scrollY = 0
		List<String> filenames = []

		for (int i = 0; i < steps; i++) {
			jsExecutor.executeScript("window.scrollTo(0, arguments[0]);", scrollY)
			jsExecutor.executeScript("window.scrollBy(0, -70);")
			Thread.sleep(500)

			File screenshot = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE)
			BufferedImage image = ImageIO.read(screenshot)

			String filename = "${NumberCapture}. ${baseFilename}_${i + 1}.png"
			File outputFile = new File("${captureDirPath}/${filename}")
			ImageIO.write(image, "png", outputFile)

			filenames << outputDirectory+"/"+filename
			scrollY += viewportHeight - 50
		}

		return filenames
	}
}
