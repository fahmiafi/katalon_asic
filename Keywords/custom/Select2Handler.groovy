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
import com.kms.katalon.core.webui.driver.DriverFactory
import org.openqa.selenium.By
import org.openqa.selenium.WebDriver
import org.openqa.selenium.WebElement
import org.openqa.selenium.JavascriptExecutor

import internal.GlobalVariable
import logger.TestStepLogger

class Select2Handler {
	static void handleSelect2DropdownWithScreenshot(String NoTC, String stepName, int numberCaptureStart, String dirCapture, String optionTextToSelect) {
		String newDirectoryPath = GlobalVariable.PathCapture+"\\"+NoTC+"\\"+stepName
		WebDriver driver = DriverFactory.getWebDriver()
		List<WebElement> approverOptions = driver.findElements(By.xpath("//li[contains(@class,'select2-results__option')]"))
		int totalOptions = approverOptions.size()
		int numberCapture = numberCaptureStart

		if (totalOptions > 0) {
			println("Jumlah total opsi dalam Select2: " + totalOptions)

			int displayedOptions = 6
			int screenshotCount = 1

			// Ambil screenshot pertama (awal)
			List<String> imageFiles = []
			String filename = numberCapture+". Dropdown Next Approver Page_${screenshotCount}"
			WebUI.takeScreenshot(newDirectoryPath + "/${filename}.png")
			imageFiles << stepName+"/"+filename
//			numberCapture++
			println("Screenshot awal (tanpa scroll) berhasil disimpan.")

			// Scroll dan screenshot jika perlu
			if (totalOptions > displayedOptions) {
				JavascriptExecutor js = (JavascriptExecutor) driver

				for (int j = displayedOptions; j < totalOptions; j += displayedOptions) {
					WebElement nextOption = approverOptions[j]
					js.executeScript("arguments[0].scrollIntoView(true);", nextOption)
					WebUI.delay(1)
					screenshotCount++
					filename = numberCapture+". Dropdown Next Approver Page_${screenshotCount}"
					WebUI.takeScreenshot(newDirectoryPath + "/${filename}.png")
//					numberCapture++
					imageFiles << stepName+"/"+filename
					println("Screenshot halaman ke-${screenshotCount} berhasil disimpan.")
				}
			}
			WebUI.delay(1)

			// Pilih opsi berdasarkan teks
			List<WebElement> approverOption = driver.findElements(By.xpath("//li[contains(@class,'select2-results__option') and contains(text(),'" + optionTextToSelect + "')]"))
			if (approverOption.size() > 0) {
				approverOption[0].click()
				println("Klik opsi '${optionTextToSelect}' berhasil.")
			} else {
				println("Opsi '${optionTextToSelect}' tidak ditemukan.")
			}
			println(imageFiles)
			String[] imageFilesArray = imageFiles.toArray(new String[0])
			TestStepLogger.addStepWithUserAndWithOutCapture(NoTC, stepName, "Dropdown Next Approver", imageFilesArray)
		} else {
			println("Tidak ada opsi yang ditemukan dalam Select2.")
		}		
	}
}