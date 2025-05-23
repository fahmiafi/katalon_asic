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
import org.openqa.selenium.interactions.Actions
import org.openqa.selenium.Keys
import com.kms.katalon.core.testobject.ConditionType
import org.openqa.selenium.JavascriptExecutor
import org.openqa.selenium.WebElement
import org.openqa.selenium.By



newDirectoryPath = GlobalVariable.PathCapture+"\\"+NoTC+"\\"+stepName+"\\"+ApproverCount+"\\Activity-"+NumberAct
GlobalVariable.newDirectoryPath = newDirectoryPath
CustomKeywords.'custom.CustomKeywords.captureFullPageInSections'('D:/katalon/ss/ASIC/TES/', '1. Update Activity Inquired')
//
//CustomKeywords.'custom.CustomKeywords.captureFullPageInSections'('D:/katalon/ss/ASIC/TES/', '. Status Activity pada Maker Approved')
//WebUI.delay(2)
//WebUI.click(findTestObject('Object Repository/COP/div_Approval History'))
//WebUI.scrollToElement(findTestObject('Object Repository/COP/div_Approval History'), 30)
//WebUI.delay(2)
//WebUI.takeScreenshot('D:/katalon/ss/ASIC/TES/Approval History.png')
//WebUI.delay(2)