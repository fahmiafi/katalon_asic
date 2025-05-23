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
import org.openqa.selenium.support.ui.Select

import utils.LogHelper

import javax.swing.JOptionPane
import com.kms.katalon.core.util.KeywordUtil

String testCaseName = RunConfiguration.getExecutionSourceName()

// Path ke file Excel
String excelFilePath = RunConfiguration.getProjectDir() + "/Data Files/SkenarioEnhanceRekSingleSide.xlsx"
FileInputStream file = new FileInputStream(excelFilePath)
Workbook workbook = new XSSFWorkbook(file)
Sheet sheet1 = workbook.getSheet("Batch")

WebUI.openBrowser('')
WebUI.navigateToUrl('http://192.168.174.45/Login')
WebUI.maximizeWindow()

// Loop melalui data di Sheet1
for (int i = 1; i <= sheet1.getLastRowNum(); i++) {
	println("cek data ke-"+i)
	Row row = sheet1.getRow(i)
	String checkRunning = row.getCell(2).getStringCellValue()
	if (row != null && checkRunning != "") {
		String NoTC = row.getCell(0).getStringCellValue()
		String NoMemo = row.getCell(1).getStringCellValue()
		String IsRunning = row.getCell(2).getStringCellValue()
		String UseCase = row.getCell(3).getStringCellValue()
		String Segmen = row.getCell(4).getStringCellValue()
		String Pencairan = row.getCell(5).getStringCellValue()
		String Skenario = row.getCell(6).getStringCellValue()
		String Nominal = String.valueOf((long) row.getCell(7).getNumericCellValue())
		String RMNpp = String.valueOf((long) row.getCell(8).getNumericCellValue())
		String RMName = row.getCell(9).getStringCellValue()
		String MakerNpp = String.valueOf((long) row.getCell(10).getNumericCellValue())
		String MakerPassword = row.getCell(11).getStringCellValue()
		String MakerName = row.getCell(12).getStringCellValue()
		String MakerPositionName = row.getCell(13).getStringCellValue()
		String MakerRole = row.getCell(14).getStringCellValue()
		String NextApproverNpp_1 = String.valueOf((long) row.getCell(15).getNumericCellValue())
		String NextApproverPassword_1 = row.getCell(16).getStringCellValue()
		String NextApproverName_1 = row.getCell(17).getStringCellValue()
		String NextApproverPositionName_1 = row.getCell(18).getStringCellValue()
		String NextApproverRole_1 = row.getCell(19).getStringCellValue()
		String NextApproverNpp_2 = String.valueOf((long) row.getCell(20).getNumericCellValue())
		String NextApproverPassword_2 = row.getCell(21).getStringCellValue()
		String NextApproverName_2 = row.getCell(22).getStringCellValue()
		String NextApproverPositionName_2 = row.getCell(23).getStringCellValue()
		String NextApproverRole_2 = row.getCell(24).getStringCellValue()
		String NextApproverNpp_3 = String.valueOf((long) row.getCell(25).getNumericCellValue())
		String NextApproverPassword_3 = row.getCell(26).getStringCellValue()
		String NextApproverName_3 = row.getCell(27).getStringCellValue()
		String NextApproverPositionName_3 = row.getCell(28).getStringCellValue()
		String NextApproverRole_3 = row.getCell(29).getStringCellValue()
		String NextApproverNpp_4 = String.valueOf((long) row.getCell(30).getNumericCellValue())
		String NextApproverPassword_4 = row.getCell(31).getStringCellValue()
		String NextApproverName_4 = row.getCell(32).getStringCellValue()
		String NextApproverPositionName_4 = row.getCell(33).getStringCellValue()
		String NextApproverRole_4 = row.getCell(34).getStringCellValue()
		String NextApproverNpp_5 = String.valueOf((long) row.getCell(35).getNumericCellValue())
		String NextApproverPassword_5 = row.getCell(36).getStringCellValue()
		String NextApproverName_5 = row.getCell(37).getStringCellValue()
		String NextApproverPositionName_5 = row.getCell(38).getStringCellValue()
		String NextApproverRole_5 = row.getCell(39).getStringCellValue()
		
		String newDirectoryPath = GlobalVariable.PathCapture+"\\"+NoTC+"\\"+testCaseName
		GlobalVariable.newDirectoryPath = newDirectoryPath
		Integer numberCapture = 1
		
		File directory = new File(newDirectoryPath)
		directory.mkdirs()
		
		String[][] nppAndNamaApproval = [
			[NextApproverNpp_1, NextApproverPassword_1, NextApproverName_1],
			[NextApproverNpp_2, NextApproverPassword_2, NextApproverName_2],
			[NextApproverNpp_3, NextApproverPassword_3, NextApproverName_3],
			[NextApproverNpp_4, NextApproverPassword_4, NextApproverName_4],
			[NextApproverNpp_5, NextApproverPassword_5, NextApproverName_5]
		]
		
		Integer TotalApproval = 1;
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
					TotalApproval = TotalApproval+1;
					NextApprover = nppAndNamaApproval[IndexNextApproval][2]
				}
			}
			
			if (NextApprover == "") {
				break
			}
		}
		
		WebDriver driver = DriverFactory.getWebDriver()
		// Contoh akses data
		println (">>>>>>>>>"+NoTC+' '+ UseCase + ' '+ Skenario +' - '+ Pencairan+"<<<<<<<<<<")
		
		// Tampilkan konfirmasi pop-up
		int result = JOptionPane.showConfirmDialog(null,
			"Silakan update data di database terlebih dahulu.\nKlik 'Yes' jika sudah selesai.",
			"Konfirmasi",
			JOptionPane.YES_NO_OPTION)
		
		if (result != JOptionPane.YES_OPTION) {
			KeywordUtil.logInfo("Eksekusi dibatalkan oleh user.")
			assert false // Menghentikan eksekusi jika user menekan 'No'
		}
		
		KeywordUtil.logInfo("Melanjutkan eksekusi setelah konfirmasi...")
		
		Integer ApproverCount = 1
		String textLog = "PASS"
		LogHelper.writeLog(testCaseName, NoTC+" "+Segmen+" "+UseCase, "START")
		for (int j = 0; j < nppAndNamaApproval.length; j++) {
//		for (int j = 1; j < nppAndNamaApproval.length; j++) {
			String nppApproval = nppAndNamaApproval[j][0]
			String PasswordApproval = nppAndNamaApproval[j][1]
			String NamaApproval = nppAndNamaApproval[j][2]
			Integer IndexNextApproval = 0
			Boolean isPerubahanData_Pinbuk = false
			String NextApprover = ''
			if (j < nppAndNamaApproval.length) {
				IndexNextApproval = j + 1
				if (nppAndNamaApproval.length == IndexNextApproval) {
					NextApprover = ""
				}
				else {
					if (Segmen == 'BOP') {
						NextApprover = nppAndNamaApproval[IndexNextApproval][0]
					}
					else {
						NextApprover = nppAndNamaApproval[IndexNextApproval][2]						
					}
				}
			}
			
			if (UseCase == 'Pemindahbukuan') {
				Sheet sheet3 = workbook.getSheet("PinbukNonPenc")
				String SelectBucket = ""
				String RekGL = ""
				String NoRekSingleSide = ""
				for (int k = 1; k <= sheet3.getLastRowNum(); k++) {
					Row rowSheetUseCase = sheet3.getRow(k)
					if (rowSheetUseCase != null && rowSheetUseCase.getCell(0).getStringCellValue() == NoTC) {
						SelectBucket = rowSheetUseCase.getCell(8).getStringCellValue()
						RekGL = rowSheetUseCase.getCell(9).getStringCellValue()
						NoRekSingleSide = rowSheetUseCase.getCell(10).getStringCellValue()
						break
					}
				}
				
				if (RekGL != '' || NoRekSingleSide != '') {
					isPerubahanData_Pinbuk = true
				}
			}
			
			String CommentApprove = 'oke setuju'
			Boolean RejectBatch = false
			
			// Login Approval
			WebUI.setText(findTestObject('Object Repository/Login/inputtxtUsername'), nppApproval)
			WebUI.setText(findTestObject('Object Repository/Login/inputtxtPassword'), PasswordApproval)
			WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Login - Approval '+ApproverCount+'.png')
			WebUI.click(findTestObject('Object Repository/Login/button_Sign In'))
			
			if (Segmen == 'BOP') {
				WebUI.waitForElementVisible(findTestObject('Object Repository/BOP/CreateNewBatch/p_Admin Kredit_Menu'), 30)
				// View Batch
				WebUI.click(findTestObject('Object Repository/BOP/CreateNewBatch/p_Admin Kredit_Menu'))
				WebUI.click(findTestObject('Object Repository/BOP/a_Approval Admin Kredit_SubMenu'))
			}
			else {
				WebUI.waitForElementVisible(findTestObject('Object Repository/COP/a_Admin Kredit'), 30)
				// View Batch
				WebUI.click(findTestObject('Object Repository/COP/a_Admin Kredit'))
				WebUI.click(findTestObject('Object Repository/COP/Approval_Object/a_Approval Cop'))
			}
			WebUI.delay(2)
			
			String targetName = NoTC
			println ("Cari : "+ targetName)
			String tableXPath = "//table[contains(@class, 'jsgrid-table')]"
			List<WebElement> rows = driver.findElements(By.xpath(tableXPath + "/tbody/tr"))
			
			int rowIndex = -1
			for (int k = 0; k < rows.size(); k++) {
				WebElement cell = rows[k].findElement(By.xpath("./td[2]"))
//				if (cell.getText().trim().equals(targetName)) {
				if (cell.getText().trim().toLowerCase().contains(targetName.toLowerCase())) {
					rowIndex = k + 1
					println("Baris ditemukan di indeks: " + rowIndex)
					((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView({block: 'center'});", cell)
					cell.click()
					WebUI.delay(1)
					WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Data Batch Approval - Approval '+ApproverCount+'.png')
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
			WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Page Batch - Approval '+ApproverCount+'.png')
			WebUI.scrollToElement(findTestObject('Object Repository/COP/Approval_Object/label_Appoval_List Aktivitas'), 30)
			WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. View Activity - Approval '+ApproverCount+'.png')
			
			// View Activity
			WebUI.click(findTestObject('Object Repository/COP/Approval_Object/button_Approval_View'))
			WebUI.delay(1)
			println("isPerubahanData_Pinbuk: "+isPerubahanData_Pinbuk)
			if (isPerubahanData_Pinbuk == true) {
				println("terdapat perubahan data")
				WebUI.click(findTestObject('Object Repository/Activity/ActivtyPemindahbukuan/div_Pinbuk_Perubahan Data'))
			}
			CustomKeywords.'custom.CustomKeywords.captureFullPageInSections'(newDirectoryPath+'/', numberCapture++ +'. Form Approval Confirm - Approval '+ApproverCount)
			
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
//					println("Dropdown dengan ID '${dropdownId}' tidak ditemukan atau terjadi error: ${e.getMessage()}")
					println("Dropdown dengan ID '${dropdownId}' tidak ditemukan")
				}
			}
			
			// Confirm Activity
			WebUI.click(findTestObject('Object Repository/COP/Approval_Object/button_Approval_Confirm'))
			WebUI.waitForElementVisible(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_OK_sukses_submit'), 30)
			WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Sukses Confirm - Approval '+ApproverCount+'.png')
			WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_OK_sukses_submit'))
			
			WebUI.scrollToElement(findTestObject('Object Repository/COP/Approval_Object/label_Appoval_List Aktivitas'), 30)
			WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Activity After Confirm - Approval '+ApproverCount+'.png')
			
			// Submit to Next Approver
			JavascriptExecutor jsExecutor = (JavascriptExecutor) driver
			jsExecutor.executeScript("window.scrollTo(0, document.body.scrollHeight);")
			
			String statusActivity = WebUI.getText(findTestObject('Object Repository/BOP/td_StatusActivity'))
			println(statusActivity)
			if(statusActivity.contains("Approved") && Segmen == 'BOP') {
				NextApprover = ''
			}
			
			println("Next Approver: "+NextApprover)
			
			if (NextApprover != '' && NextApprover != '0') {
				if (Segmen == 'BOP') {
					WebUI.click(findTestObject('Object Repository/BOP/select_NextApprover_BOP'))
					WebUI.delay(2)
					WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +". Dropdown Next Approver Page_1.png")
					WebElement dropdown = driver.findElement(By.xpath('//select[@id="NextApprover"]')) // Sesuaikan dengan XPath elemen
					Select select = new Select(dropdown)
					select.selectByValue(NextApprover)
				}
				else {
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
							List<WebElement> approverOptions = driver.findElements(By.xpath("//li[contains(@class,'select2-results__option')]"))
							// Pastikan ada opsi yang ditemukan
							int totalOptions = approverOptions.size()
							if (totalOptions > 0) {
								println("Jumlah total opsi dalam Select2: " + totalOptions)
							
								int displayedOptions = 6  // Jumlah opsi yang ditampilkan per halaman
								int screenshotCount = 1    // Nomor urut screenshot
								
								// **1️⃣ Ambil screenshot pertama sebelum scroll**
								WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +". Dropdown Next Approver Page_${screenshotCount} - Approval ${ApproverCount}.png")
								println("Screenshot awal (tanpa scroll) berhasil disimpan.")
								
								// **2️⃣ Scroll bertahap jika opsi lebih dari 6**
								if (totalOptions > displayedOptions) {
									JavascriptExecutor js = (JavascriptExecutor) driver
									
									for (int k = displayedOptions; k < totalOptions; k += displayedOptions) {
										WebElement nextOption = approverOptions[k]
										
										// **Scroll ke opsi berikutnya**
										js.executeScript("arguments[0].scrollIntoView(true);", nextOption)
										WebUI.delay(1) // Jeda agar scroll berjalan dengan baik
										
										// **Ambil screenshot setelah scroll**
										screenshotCount++
										WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +". Dropdown Next Approver Page_${screenshotCount} - Approval ${ApproverCount}.png")
										println("Screenshot halaman ke-${screenshotCount} berhasil disimpan.")
									}
								}
								
								// pilih next approver
								List<WebElement> approverOption = driver.findElements(By.xpath("//li[contains(@class,'select2-results__option') and contains(text(),'" + NextApprover + "')]"))
								approverOption[0].click()
								println("Klik opsi ${NextApproverName_1} berhasil.")
							} else {
								println("Tidak ada opsi yang ditemukan dalam Select2.")
							}
						}
					} else {
						println("Elemen ApprovalList tidak ditemukan, Batch di Reject")
						RejectBatch = true
					}
				}
			}
			
			if (RejectBatch == true) {
				CommentApprove = 'Reject, Approval selanjutnya tidak ada'
				NextApprover = ''
				textLog = "ERROR Approval ke-"+ApproverCount+" tidak ada"
				
				WebUI.takeScreenshot(newDirectoryPath + '/ERROR '+ NoTC +' List Approval tidak ada '+ApproverCount+'.png')
			}
			else {
				WebUI.setText(findTestObject('Object Repository/COP/Approval_Object/textarea_Approval_Comment'), CommentApprove)
				WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Submit Batch - Approval '+ApproverCount+'.png')
				
				// Submit Batch
				WebUI.click(findTestObject('Object Repository/COP/Approval_Object/button_Submit_Batch_Approval'))
				WebUI.waitForElementVisible(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_OK_sukses_submit'), 30)
				WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Sukses Submit - Approval '+ApproverCount+'.png')
				WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_OK_sukses_submit'))
				
			}
			
			WebUI.delay(2)
			// Logout
			WebUI.click(findTestObject('Object Repository/Login/i_User Logout'))
			WebUI.click(findTestObject('Object Repository/Login/a_Logout'))
			
			WebUI.delay(1)
			LogHelper.writeLog(testCaseName, NoTC+" "+Segmen+" "+UseCase, "Approver "+ApproverCount+"/"+ TotalApproval + " Sukses || "+nppApproval+ " "+NamaApproval)
			ApproverCount++
			
			if (NextApprover == '') {
				break
			}
		}
		
		// Cek status pada user maker
		LogHelper.writeLog(testCaseName, NoTC+" "+Segmen+" "+UseCase, "Cek status pada user maker")
		WebUI.setText(findTestObject('Object Repository/Login/inputtxtUsername'), MakerNpp)
		WebUI.setText(findTestObject('Object Repository/Login/inputtxtPassword'), MakerPassword)
		WebUI.click(findTestObject('Object Repository/Login/button_Sign In'))
		if (Segmen == 'BOP') {
			WebUI.waitForElementVisible(findTestObject('Object Repository/BOP/CreateNewBatch/p_Admin Kredit_Menu'), 30)
			// View Batch
			WebUI.click(findTestObject('Object Repository/BOP/CreateNewBatch/p_Admin Kredit_Menu'))
			WebUI.click(findTestObject('Object Repository/BOP/CreateNewBatch/a_Admin Kredit_SubMenu'))
		}
		else {
			WebUI.waitForElementVisible(findTestObject('Object Repository/COP/a_Admin Kredit'), 30)
			// View Batch
			WebUI.click(findTestObject('Object Repository/COP/a_Admin Kredit'))
			WebUI.click(findTestObject('Object Repository/COP/a_Monitoring Batch Progress  Failed'))
		}
		WebUI.waitForElementVisible(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/search_button'), 30)
		WebUI.setText(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/input_filter_no_batch'), NoMemo)
		WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/search_button'))
		
		// Klik tombol View jika ditemukan
		WebUI.waitForElementVisible(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_View'), 30)
		WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Search Batch.png')
		WebUI.click(findTestObject('Object Repository/COP/UpdateAfterInquiry_Object/button_View'))
		
		CustomKeywords.'custom.CustomKeywords.captureFullPageInSections'(newDirectoryPath+'/', numberCapture++ +'. Batch Full')
		WebUI.click(findTestObject('Object Repository/COP/div_Approval History'))
		WebUI.scrollToElement(findTestObject('Object Repository/COP/div_Approval History'), 30)
//		WebUI.scrollToElement(findTestObject('Object Repository/COP/Page_BNI.RPA.CORE/button_View Summary'), 30)
		WebUI.takeScreenshot(newDirectoryPath + '/'+ numberCapture++ +'. Approval History.png')
		WebUI.delay(2)
		// Logout
		WebUI.click(findTestObject('Object Repository/Login/i_User Logout'))
		WebUI.click(findTestObject('Object Repository/Login/a_Logout'))
		
		// tulis log
		LogHelper.writeLog(testCaseName, NoTC+" "+Segmen+" "+UseCase, "END")
	}
	else {
		println("data ke:"+i+" tidak dijalankan")
	}
}

// Tutup
workbook.close()
file.close()
WebUI.closeBrowser()