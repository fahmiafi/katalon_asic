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
import org.openqa.selenium.JavascriptExecutor
import org.openqa.selenium.By
import org.openqa.selenium.WebElement
import com.kms.katalon.core.testobject.ConditionType
import javax.swing.JOptionPane
import utils.LogHelper
import excel.ExcelHelper
import approval.ApprovalHelper
import logger.TestStepLogger
import custom.Select2Handler
import com.kms.katalon.core.util.KeywordUtil

WebUI.openBrowser('')
WebUI.navigateToUrl('http://192.168.174.45/Login')
WebUI.maximizeWindow()

TestStepLogger.addStepWithUserAndCapture("001", "Maker", 1, 1, "test", "Maker", true, true)
TestStepLogger.addStepWithUserAndCapture("001", "Maker", 1, 2, "test", "Maker", true, true)
TestStepLogger.addOutputWithUserAndCapture("001", "Maker", 1, 1, "test output", "Maker", true, true)
TestStepLogger.addOutputWithUserAndCapture("001", "Maker", 2, 1, "test outourrr", "Maker", true, true)
TestStepLogger.addOutputWithUserAndCapture("001", "Maker", 3, 1, "test outpurrrrew", "Maker", true, true)
TestStepLogger.addStepWithUserAndCapture("001", "Maker", 1, 2, "test", "Maker", true, true)