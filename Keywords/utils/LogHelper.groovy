package utils

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
import java.text.SimpleDateFormat
import java.nio.file.Files
import java.nio.file.Paths

public class LogHelper {
	static void writeLog(String testCaseName, String skenario, String result) {
		String currentDate = new SimpleDateFormat("yyyy-MM-dd").format(new Date())
		String currentTime = new SimpleDateFormat("HH:mm:ss").format(new Date())

		String logDir = "Logs"
		Files.createDirectories(Paths.get(logDir))

		String logFilePath = "${logDir}/log_${testCaseName}_${currentDate}.txt"

		String logEntry = "[${currentDate} ${currentTime}] Skenario: ${skenario} | Result: ${result}\n"

		Files.write(Paths.get(logFilePath), logEntry.getBytes(), java.nio.file.StandardOpenOption.CREATE, java.nio.file.StandardOpenOption.APPEND)

		println "✅ Log berhasil disimpan di: ${logFilePath}"
	}
}
