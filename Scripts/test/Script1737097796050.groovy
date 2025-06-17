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


String stepName = "Approval"

// Path ke file Excel
String excelFilePath = RunConfiguration.getProjectDir() + GlobalVariable.PathDataExcel
FileInputStream file = new FileInputStream(excelFilePath)
Workbook workbook = new XSSFWorkbook(file)
Sheet sheetBatch = workbook.getSheet("Batch")
Sheet sheetActivity = workbook.getSheet("Activity")
Sheet sheetApproval = workbook.getSheet("Approval")
Sheet sheetPemindahbukuan = workbook.getSheet("Act Pemindahbukuan")
Sheet sheetPembukaan = workbook.getSheet("Act Pembukaan Rek")
String ApprovalId = "1"
String Segmen = "Korporasi & Enterprise"

WebUI.openBrowser('')
WebUI.navigateToUrl('http://192.168.174.45/Login')
WebUI.maximizeWindow()
for(int i; i < 5; i++) {
	println("cek ${i}")
	if(i == 3) {
		ApprovalId = null
	}
	if(i == 4) {
		ApprovalId = "7"
	}
	def resultApproval = []
	String[][] nppAndNamaApproval
	if (ApprovalId != null) {
		// Search Approval by Id
		resultApproval = ApprovalHelper.getApprovalDataById(sheetApproval, ApprovalId, Segmen)
	}
	else {
		// Search Approval by Activity
		resultApproval  = ApprovalHelper.getApprovalData(
			sheetActivity,
			sheetPemindahbukuan,
			sheetPembukaan,
			sheetApproval,
			"A003",
			Segmen
		)
	}
	
	if(resultApproval == null) {
		int alertApprovalNotFound = JOptionPane.showMessageDialog(null,
			"Data Approval tidak ditemukan, silahkan cek terlebih dahulu.\n Untuk sementara Test Case dihentikan",
			"Approval Not Found",
			JOptionPane.YES_OPTION)
		KeywordUtil.logInfo("Eksekusi dibatalkan oleh user.")
		assert false // Menghentikan eksekusi jika user menekan 'No'
	}
	else {
		def dataApproval = resultApproval.dataApproval
		def approvalIdTerbanyak = resultApproval.maxApprovalId
		nppAndNamaApproval = dataApproval
		
		println(nppAndNamaApproval)
	}
	
	Integer ApproverCount = 1
	for (int j = 0; j < nppAndNamaApproval.length; j++) {
		String nppApproval = nppAndNamaApproval[j][0]
		String PasswordApproval = nppAndNamaApproval[j][1]
		String NamaApproval = nppAndNamaApproval[j][2]
		Integer IndexNextApproval = 0
		String NextApprover = ''
		if (j < nppAndNamaApproval.length) {
			IndexNextApproval = j + 1
			if (nppAndNamaApproval.length == IndexNextApproval) {
				NextApprover = ""
			}
			else {
				NextApprover = nppAndNamaApproval[IndexNextApproval][2]
			}
		}
		
		String CommentApprove = 'oke setuju'
		Boolean RejectBatch = false
		
		// Login Approval
		WebUI.setText(findTestObject('Object Repository/Login/inputtxtUsername'), nppApproval)
		WebUI.setText(findTestObject('Object Repository/Login/inputtxtPassword'), PasswordApproval)
		WebUI.click(findTestObject('Object Repository/Login/button_Sign In'))
		
		
		WebUI.delay(2)
		// Logout
		WebUI.click(findTestObject('Object Repository/Login/i_User Logout'))
		WebUI.click(findTestObject('Object Repository/Login/a_Logout'))
		
		WebUI.delay(1)
//			LogHelper.writeLog(testCaseName, NoTC+" "+Segmen+" "+UseCase, "Approver "+ApproverCount+"/"+ TotalApproval + " Sukses || "+nppApproval+ " "+NamaApproval)
		ApproverCount++
		
		if (NextApprover == '') {
			break
		}
	}
}

