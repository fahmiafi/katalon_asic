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
import com.kms.katalon.core.webui.common.WebUiCommonHelper
import com.kms.katalon.core.testobject.ConditionType

import utils.LogHelper
import javax.swing.JOptionPane
import com.kms.katalon.core.util.KeywordUtil
import excel.ExcelHelper
import approval.ApprovalHelper
import logger.TestStepLogger
import custom.Select2Handler

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

WebUI.openBrowser('')
WebUI.navigateToUrl('http://192.168.174.45/Login')
WebUI.maximizeWindow()

// Loop melalui data di Sheet1
for (int i = 1; i <= sheetBatch.getLastRowNum(); i++) {
	Row row = sheetBatch.getRow(i)
	String checkRunning = row.getCell(2).getStringCellValue()
	if (row != null && checkRunning == "Y") {
		String NoTC = ExcelHelper.getCellValueAsString(row, 0)
		String NoMemo = ExcelHelper.getCellValueAsString(row, 1)
		String IsRunning = ExcelHelper.getCellValueAsString(row, 2)
		String Segmen = ExcelHelper.getCellValueAsString(row, 3)
		String SkenarioBatch = ExcelHelper.getCellValueAsString(row, 4)
		String RMNpp = ExcelHelper.getCellValueAsString(row, 5)
		String RMName = ExcelHelper.getCellValueAsString(row, 6)
		String MakerNpp = ExcelHelper.getCellValueAsString(row, 7)
		String MakerPassword = ExcelHelper.getCellValueAsString(row, 8)
		String MakerName = ExcelHelper.getCellValueAsString(row, 9)
		String MakerPositionName = ExcelHelper.getCellValueAsString(row, 10)
		String MakerRole = ExcelHelper.getCellValueAsString(row, 11)
		String ApprovalId = ExcelHelper.getCellValueAsString(row, 12)
		
		def resultApproval = []
		String[][] nppAndNamaApproval
		if (ApprovalId != null) {
			// Search Approval by Id
			println("Search Approval by Id")
			resultApproval = ApprovalHelper.getApprovalDataById(sheetApproval, ApprovalId, Segmen)
		}
		else {
			// Search Approval by Activity
			println("Search Approval by Activity")
			resultApproval  = ApprovalHelper.getApprovalData(
				sheetActivity,
				sheetPemindahbukuan,
				sheetPembukaan,
				sheetApproval,
				NoTC,
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
		
		WebDriver driver = DriverFactory.getWebDriver()
		
		println (">>>>>>>>>"+NoTC+' '+ SkenarioBatch +"<<<<<<<<<<")
		
		// Tampilkan konfirmasi pop-up
		int alertUpdateDB = JOptionPane.showConfirmDialog(null,
			"Silakan update data di database terlebih dahulu.\nKlik 'Yes' jika sudah selesai.",
			"Konfirmasi",
			JOptionPane.YES_NO_OPTION)
		
		if (alertUpdateDB != JOptionPane.YES_OPTION) {
			KeywordUtil.logInfo("Eksekusi dibatalkan oleh user.")
			assert false // Menghentikan eksekusi jika user menekan 'No'
		}
		
		KeywordUtil.logInfo("Melanjutkan eksekusi setelah konfirmasi...")
		
		Integer ApproverCount = 1
		String textLog = "PASS"
//		LogHelper.writeLog(testCaseName, NoTC+" "+Segmen+" "+SkenarioBatch, "START")
		for (int j = 0; j < nppAndNamaApproval.length; j++) {
			dirCapture = stepName+"-"+ApproverCount
			numberCapture = 1
			
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
			TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, numberCapture++, 'Login sebagai Approval', dirCapture, true, false)
			WebUI.click(findTestObject('Object Repository/Login/button_Sign In'))
			
			WebUI.waitForElementVisible(findTestObject('Object Repository/COP/a_Admin Kredit'), 30)
			
			// Search Batch
			TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, numberCapture, 'Akses menu Admin Kredit >> Approval COP', dirCapture, false, false)
			WebUI.click(findTestObject('Object Repository/COP/a_Admin Kredit'))
			WebUI.waitForElementVisible(findTestObject('Object Repository/COP/Approval_Object/a_Approval Cop'), 30)
			WebUI.click(findTestObject('Object Repository/COP/Approval_Object/a_Approval Cop'))
			WebUI.delay(2)
			
			String targetName = NoTC
			println ("Cari : "+ targetName)
			String tableXPath = "//table[contains(@class, 'jsgrid-table')]"
			List<WebElement> rows = driver.findElements(By.xpath(tableXPath + "/tbody/tr"))
			
			int rowIndex = -1
			for (int k = 0; k < rows.size(); k++) {
				WebElement cell = rows[k].findElement(By.xpath("./td[2]"))
				if (cell.getText().trim().toLowerCase().contains(targetName.toLowerCase())) {
					rowIndex = k + 1
					println("Baris ditemukan di indeks: " + rowIndex)
					((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView({block: 'center'});", cell)
					cell.click()
					WebUI.delay(1)
					TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, numberCapture++, 'Klik action pada data yang telah di submit oleh user Maker', dirCapture, true, false)
					break
				}
			}
			
			// Klik tombol aksi di baris yang ditemukan
			if (rowIndex != -1) {
				String buttonXPath = tableXPath + "/tbody/tr[" + rowIndex + "]/td[last()]/div/button"
				WebElement actionButton = driver.findElement(By.xpath(buttonXPath))
				actionButton.click()
				println("Tombol aksi diklik pada baris: " + rowIndex)
			} else {
				println("Nama tidak ditemukan di kolom 2.")
			}
			
			// Halaman Batch Approval
			WebUI.delay(1)
			TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, numberCapture++, 'Data batch yang akan di approve', dirCapture, true, false)
			WebUI.scrollToElement(findTestObject('Object Repository/COP/Approval_Object/label_Flag Batch'), 30)
			TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, numberCapture++, 'Klik View pada activity', dirCapture, true, false)
			
			// Ambil semua tombol View di dalam tabel
			List<WebElement> viewButtons = WebUiCommonHelper.findWebElements(
				new TestObject().tap {
					addProperty("xpath", ConditionType.EQUALS, "//table[@id='activityTable']//button[contains(., 'View')]")
				},
				10
			)
			
			println "Jumlah View buttons: ${viewButtons.size()}"
			
			int NumberAct = 1;
			for (int a = 0; a < viewButtons.size(); a++) {
				numberCaptureAct = 1
				dirCapture = stepName+"-"+ApproverCount+"/Form Approve Activity-"+NumberAct
				// Ambil ulang semua tombol View karena DOM berubah setelah navigasi
				viewButtons = WebUiCommonHelper.findWebElements(
					new TestObject().tap {
						addProperty("xpath", ConditionType.EQUALS, "//table[@id='activityTable']//button[contains(., 'View')]")
					},
					10
				)
			
				// Klik tombol View ke-i
				WebUI.executeJavaScript("arguments[0].click();", Arrays.asList(viewButtons[a]))
				WebUI.waitForPageLoad(10)
				WebUI.delay(1) // opsional: beri waktu agar halaman benar-benar siap
				TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, numberCaptureAct++, 'Approve pada Form Activity pada halaman approval', dirCapture, true, true)
			
				List<String> dropdownIds = [
					"Action",
					"ApprovalPembukaanRek_PembukaanRekeningPinjaman",
					"ApprovalPembukaanRek_PencairanDana"
				]
				
				// Iterasi melalui setiap ID dropdown
				for (String dropdownId : dropdownIds) {
					try {
						// Temukan elemen dropdown berdasarkan ID
						WebElement selectElement = driver.findElement(By.xpath("//select[@id='" + dropdownId + "']"))
						
						if (selectElement != null) {
							// Ambil semua opsi dalam dropdown
							List<WebElement> options = selectElement.findElements(By.tagName("option"))
							String selectedValue = null
				
							// Tentukan opsi yang akan dipilih berdasarkan logika
							for (WebElement option : options) {
								String value = option.getAttribute("value")
								if (value == 'Approve') {
									selectedValue = 'Approve'
									break
								} else if (value == 'Check' && selectedValue == null) {
									selectedValue = 'Check'
								}
							}
				
							// Gunakan TestObject untuk melakukan select
							TestObject dropdownApproveOrReject = new TestObject()
							dropdownApproveOrReject.addProperty("xpath", com.kms.katalon.core.testobject.ConditionType.EQUALS, "//select[@id='" + dropdownId + "']")
							
							WebUI.selectOptionByValue(dropdownApproveOrReject, selectedValue, true)
							println("Dropdown dengan ID '${dropdownId}' berhasil diproses dengan nilai '${selectedValue}'")
						}
					} catch (Exception e) {
						// Jika dropdown tidak ditemukan atau ada kesalahan lain
						println("Dropdown dengan ID '${dropdownId}' tidak ditemukan atau terjadi error: ${e.getMessage()}")
					}
				}
				
				// Confirm Activity
				WebUI.scrollToElement(findTestObject('Object Repository/COP/Approval_Object/button_Approval_Confirm'), 30)
				TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, numberCapture++, 'Confirm Activity', dirCapture, true, false)
				WebUI.click(findTestObject('Object Repository/COP/Approval_Object/button_Approval_Confirm'))
				WebUI.waitForElementVisible(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_OK_sukses_submit'), 30)
				TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, numberCapture++, 'Confirm berhasil', dirCapture, true, false)
				WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_OK_sukses_submit'))
			
				WebUI.delay(2)
				WebUI.scrollToElement(findTestObject('Object Repository/COP/Approval_Object/label_Flag Batch'), 30)
				NumberAct++
			}
			dirCaptue = stepName+"-"+ApproverCount
			
			TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, numberCapture++, 'Status activity Check atau Approved', dirCapture, true, false)
			
			// Submit to Next Approver
			JavascriptExecutor jsExecutor = (JavascriptExecutor) driver
			jsExecutor.executeScript("window.scrollTo(0, document.body.scrollHeight);")
			
			println("Next Approver: "+NextApprover)
			
			if (NextApprover != '') {
				WebElement isApprovalListPresent = driver.findElement(By.id("ApproverDropdown"))
				if (isApprovalListPresent != null) {
					String style = isApprovalListPresent.getAttribute("style")
					if (style.contains("display: none;")) {
						println("Elemen ApprovalList ada tetapi disembunyikan.")
						RejectBatch = true
					} else {
						println("Elemen ada dan terlihat.")
						WebUI.click(findTestObject('Object Repository/COP/Approval_Object/span_Approval_list_option'))
						WebUI.delay(1)
						Select2Handler.handleSelect2DropdownWithScreenshot(NoTC, stepName, numberCapture++, dirCapture, NextApprover)
					}
				} else {
					println("Elemen ApprovalList tidak ditemukan, Batch di Reject")
					RejectBatch = true
				}
			}
			
			if (RejectBatch == true) {
				CommentApprove = 'Reject, Approval selanjutnya tidak ada'
				NextApprover = ''
				textLog = "ERROR Approval ke-"+ApproverCount+" tidak ada"
				
				TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, numberCapture++, 'ERROR '+ NoTC +' List Approval tidak ada'+ApproverCount, dirCapture, true, false)
				
			}
			else {
				if(NextApprover != "") {
					CommentApprove = CommentApprove+" Check"
				}
				else {
					CommentApprove = CommentApprove+" Approve"
				}
				
				WebUI.scrollToElement(findTestObject('Object Repository/COP/Approval_Object/button_Submit_Batch_Approval'), 30)
				WebUI.setText(findTestObject('Object Repository/COP/Approval_Object/textarea_Approval_Comment'), CommentApprove)
				TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, numberCapture++, 'Pilih Next Approver selanjutnya, input Comment dan Submit Batch', dirCapture, true, false)
				
				// Submit Batch
				WebUI.click(findTestObject('Object Repository/COP/Approval_Object/button_Submit_Batch_Approval'))
				WebUI.waitForElementVisible(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_OK_sukses_submit'), 30)
				TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, numberCapture++, 'Submit Batch berhasil', dirCapture, true, false)
				WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_OK_sukses_submit'))
				
			}
			
			WebUI.delay(2)
			// Logout
			WebUI.click(findTestObject('Object Repository/Login/i_User Logout'))
			WebUI.click(findTestObject('Object Repository/Login/a_Logout'))
			
			WebUI.delay(1)
			ApproverCount++
			
			if (NextApprover == '') {
				WebUI.setText(findTestObject('Object Repository/Login/inputtxtUsername'), MakerNpp)
				WebUI.setText(findTestObject('Object Repository/Login/inputtxtPassword'), MakerPassword)
				WebUI.click(findTestObject('Object Repository/Login/button_Sign In'))
				
				WebUI.waitForElementVisible(findTestObject('Object Repository/COP/a_Admin Kredit'), 30)
				
				WebUI.click(findTestObject('Object Repository/COP/a_Admin Kredit'))
				WebUI.click(findTestObject('Object Repository/COP/a_Monitoring Batch Progress  Failed'))
				WebUI.waitForElementVisible(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/search_button'), 30)
				WebUI.setText(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/input_filter_no_batch'), NoMemo)
				WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/search_button'))
				
				// Klik tombol View jika ditemukan
				WebUI.delay(3)
				WebUI.waitForElementVisible(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_View'), 30)
				TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, numberCapture++, 'Status Batch pada Maker Approved', dirCapture, true, false)
				WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_View'))
				
				TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, numberCapture++, 'Status Activity pada Maker Approved', dirCapture, true, true)
				WebUI.click(findTestObject('Object Repository/COP/div_Approval History'))
				WebUI.delay(2)
				WebUI.scrollToElement(findTestObject('Object Repository/COP/div_Approval History'), 30)
				WebUI.delay(1)
				TestStepLogger.addStepWithUserAndCapture(NoTC, stepName, numberCapture++, 'Approval History', dirCapture, true, false)
				WebUI.delay(2)
				// Logout
				WebUI.click(findTestObject('Object Repository/Login/i_User Logout'))
				WebUI.click(findTestObject('Object Repository/Login/a_Logout'))
				break
			}
		}
	}
	else {
		println("data ke:"+i+" tidak dijalankan")
	}
}

// Tutup
workbook.close()
file.close()
WebUI.closeBrowser()