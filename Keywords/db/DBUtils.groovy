package db

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

import java.sql.*

public class DBUtils {
	private static Connection connection = null


	static def connectDB() {
		if (connection != null && !connection.isClosed()) {
			return
		}

		try {
			String dbUrl = 'jdbc:sqlserver://'+GlobalVariable.sqlserverHost+';databaseName='+GlobalVariable.sqlserverDBName+';encrypt=false;trustServerCertificate=true'
			String dbUser = GlobalVariable.sqlserverUsername
			String dbPass = GlobalVariable.sqlserverPassword
			// Load SQL Server JDBC driver
			Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver")
			connection = DriverManager.getConnection(dbUrl, dbUser, dbPass)
			println("Connected to MSSQL DB")
		} catch (Exception e) {
			e.printStackTrace()
		}
	}

	static def executeUpdate(String query) {
		try {
			Statement stmt = connection.createStatement()
			int count = stmt.executeUpdate(query)
			println("${count} rows affected.")
			stmt.close()
		} catch (Exception e) {
			e.printStackTrace()
		}
	}

	static def closeConnection() {
		if (connection != null && !connection.isClosed()) {
			connection.close()
			println("Connection closed.")
		}
	}
}

