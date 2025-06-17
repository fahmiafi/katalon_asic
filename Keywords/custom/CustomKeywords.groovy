package custom

import com.kms.katalon.core.webui.driver.DriverFactory
import com.kms.katalon.core.webui.common.WebUiCommonHelper
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import org.openqa.selenium.JavascriptExecutor
import org.openqa.selenium.OutputType
import org.openqa.selenium.TakesScreenshot
import org.openqa.selenium.WebDriver
import org.openqa.selenium.io.FileHandler
import java.text.SimpleDateFormat
import java.util.Date
import java.io.File
import internal.GlobalVariable as GlobalVariable
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import org.openqa.selenium.WebElement
import java.awt.image.BufferedImage
//import java.io.File
import javax.imageio.ImageIO

import com.kms.katalon.core.testobject.TestObject as TestObject

class CustomKeywords {
	static void scrollToElement(String obj) {
		WebDriver driver = DriverFactory.getWebDriver()
		TestObject testObject = findTestObject(obj)
		WebUI.verifyElementPresent(testObject, 30)
		WebElement element = WebUI.findWebElement(testObject, 30)
		((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView({block: 'center'});", element)
	}

	static void takeScreenshot(String name) {
		String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date())
		String screenshotPath = GlobalVariable.PathCapture + name + '_' + timestamp + '.png'
		WebDriver driver = DriverFactory.getWebDriver()
		File screenshot = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE)
		FileHandler.copy(screenshot, new File(screenshotPath))
		WebUI.comment('Screenshot saved to: ' + screenshotPath)
	}

	static void captureFullPageInSections(String outputDirectory, filename) {
		WebDriver driver = DriverFactory.getWebDriver()
		JavascriptExecutor jsExecutor = (JavascriptExecutor) driver

		Long totalHeight = (Long) jsExecutor.executeScript("return document.body.scrollHeight")
		Long viewportHeight = (Long) jsExecutor.executeScript("return window.innerHeight")

		int steps = (int) Math.ceil(totalHeight / (double) viewportHeight)
		println "totalHeight: "+totalHeight+"; viewportHeight:"+viewportHeight+"; steps: "+steps
		int scrollY = 0

		for (int i = 0; i < steps; i++) {
			jsExecutor.executeScript("window.scrollTo(0, arguments[0]);", scrollY)
			jsExecutor.executeScript("window.scrollBy(0, -70);")
			Thread.sleep(500)

			File screenshot = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE)
			BufferedImage image = ImageIO.read(screenshot)

			File outputFile = new File(outputDirectory + "/"+filename+"_" + (i + 1) + ".png")
			ImageIO.write(image, "png", outputFile)

			println "Screenshot bagian " + (i + 1) + " disimpan di: " + outputFile.getAbsolutePath()

			scrollY += viewportHeight-50
		}
		jsExecutor.executeScript("window.scrollTo(0, document.body.scrollHeight);")
		File screenshot = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE)
		//		BufferedImage image = ImageIO.read(screenshot)
		//		File outputFile = new File(outputDirectory + "/"+filename+"_save.png")
		//		ImageIO.write(image, "png", outputFile)

		println "Semua screenshot bagian telah disimpan di folder: " + outputDirectory
	}

	def scrollToTop() {
		WebDriver driver = DriverFactory.getWebDriver()
		JavascriptExecutor js = (JavascriptExecutor) driver
		js.executeScript("window.scrollTo(0, 0);")
	}
}
