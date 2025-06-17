package approval

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
import excel.ExcelHelper
import org.apache.poi.ss.usermodel.*

public class ApprovalHelper {
	def static Map<String, List<List<String>>> getApprovalData(
			Sheet sheetActivity,
			Sheet sheetPemindahbukuan,
			Sheet sheetPembukaan,
			Sheet sheetApproval,
			String NoTC,
			String Segmen
	) {
		def dataApproval = [:]

		for (int a = 1; a <= sheetActivity.getLastRowNum(); a++) {
			Row rowActivity = sheetActivity.getRow(a)
			String checkTcAct = ExcelHelper.getCellValueAsString(rowActivity, 0)

			if (rowActivity != null && checkTcAct == NoTC) {
				String Seq = ExcelHelper.getCellValueAsString(rowActivity, 2)
				String UseCase = ExcelHelper.getCellValueAsString(rowActivity, 3)
				String SkenarioActivity = ExcelHelper.getCellValueAsString(rowActivity, 4)
				String Pencairan = ExcelHelper.getCellValueAsString(rowActivity, 5)

				def NominalPencairan = 0L
				if (Pencairan != "Non Pencairan") {
					if (UseCase == "Pemindahbukuan") {
						for (int b = 2; b <= sheetPemindahbukuan.getLastRowNum(); b++) {
							Row rowPemindahbukuan = sheetPemindahbukuan.getRow(b)
							if (rowPemindahbukuan != null &&
									ExcelHelper.getCellValueAsString(rowPemindahbukuan, 0) == NoTC &&
									ExcelHelper.getCellValueAsString(rowPemindahbukuan, 1) == Seq) {
								NominalPencairan = (long) rowPemindahbukuan.getCell(6).getNumericCellValue()
								break
							}
						}
					} else if (UseCase == "Pembukaan Rek") {
						for (int b = 2; b <= sheetPembukaan.getLastRowNum(); b++) {
							Row rowPembukaan = sheetPembukaan.getRow(b)
							if (rowPembukaan != null &&
									ExcelHelper.getCellValueAsString(rowPembukaan, 0) == NoTC &&
									ExcelHelper.getCellValueAsString(rowPembukaan, 1) == Seq) {
								NominalPencairan = (long) rowPembukaan.getCell(6).getNumericCellValue()
								break
							}
						}
					}
				}

				for (int b = 1; b <= sheetApproval.getLastRowNum(); b++) {
					Row rowApproval = sheetApproval.getRow(b)
					String Approval_Id = ExcelHelper.getCellValueAsString(rowApproval, 0)
					String Approval_Segmen = ExcelHelper.getCellValueAsString(rowApproval, 1)
					String Approval_Pencairan = ExcelHelper.getCellValueAsString(rowApproval, 2)
					String strNominalAkhir = ExcelHelper.getCellValueAsString(rowApproval, 5)

					boolean isMatch = false

					if (rowApproval != null && Approval_Segmen == Segmen && Approval_Pencairan == Pencairan) {
						if (Approval_Pencairan.equalsIgnoreCase("Non Pencairan")) {
							isMatch = true
						} else {
							long nominalAwal = (long) rowApproval.getCell(4).getNumericCellValue()
							if (strNominalAkhir.equalsIgnoreCase("BPMK")) {
								if (NominalPencairan >= nominalAwal) {
									isMatch = true
								}
							} else {
								long nominalAkhir = (long) rowApproval.getCell(5).getNumericCellValue()
								if (NominalPencairan >= nominalAwal && NominalPencairan <= nominalAkhir) {
									isMatch = true
								}
							}
						}

						if (isMatch) {
							def approvalList = []

							def approvers = [
								[6, 7, 8],
								[11, 12, 13],
								[16, 17, 18],
								[21, 22, 23],
								[26, 27, 28]
							]

							for (appr in approvers) {
								String npp = ExcelHelper.getCellValueAsString(rowApproval, appr[0])
								String pwd = ExcelHelper.getCellValueAsString(rowApproval, appr[1])
								String name = ExcelHelper.getCellValueAsString(rowApproval, appr[2])

								if (npp && pwd && name) {
									approvalList.add([npp, pwd, name])
								}
							}

							if (!approvalList.isEmpty()) {
								dataApproval[Approval_Id] = approvalList
							}

							break
						}
					}
				}
			}
		}

		// Ambil Approval_Id dengan jumlah approver terbanyak
		def maxKey = null
		def maxSize = 0

		dataApproval.each { key, approvers ->
			if (approvers.size() > maxSize) {
				maxSize = approvers.size()
				maxKey = key
			}
		}

		return [dataApproval: dataApproval[maxKey], maxApprovalId: maxKey]
	}
	
	def static Map<String, List<List<String>>> getApprovalDataById(Sheet sheetApproval, String id, String Segmen) {
		def dataApproval = [:]
		for (int b = 1; b <= sheetApproval.getLastRowNum(); b++) {
			Row rowApproval = sheetApproval.getRow(b)
			String Approval_Id = ExcelHelper.getCellValueAsString(rowApproval, 0)
			String Approval_Segmen = ExcelHelper.getCellValueAsString(rowApproval, 1)

			if (Approval_Id == id && Approval_Segmen == Segmen) {
				def approvalList = []
				
				def approvers = [
					[6, 7, 8],
					[11, 12, 13],
					[16, 17, 18],
					[21, 22, 23],
					[26, 27, 28]
				]
				
				for (appr in approvers) {
					String npp = ExcelHelper.getCellValueAsString(rowApproval, appr[0])
					String pwd = ExcelHelper.getCellValueAsString(rowApproval, appr[1])
					String name = ExcelHelper.getCellValueAsString(rowApproval, appr[2])
				
					if (npp && pwd && name) {
						approvalList.add([npp, pwd, name])
					}
				}
				
				if (!approvalList.isEmpty()) {
					dataApproval[Approval_Id] = approvalList
				}
				return [dataApproval: dataApproval[Approval_Id], maxApprovalId: Approval_Id]
				break
			}
		}
	}
}
