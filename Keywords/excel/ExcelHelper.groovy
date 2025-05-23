package excel

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
import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.usermodel.DateUtil

public class ExcelHelper {
    static String getCellValueAsString(Row row, int cellIndex) {
        Cell cell = row.getCell(cellIndex)

        if (cell == null) {
            return null
        }

        switch (cell.getCellTypeEnum()) {
            case CellType.STRING:
                return cell.getStringCellValue().trim()

            case CellType.NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    Date date = cell.getDateCellValue()
                    return new java.text.SimpleDateFormat("dd/MM/yyyy").format(date)
                } else {
                    return String.valueOf((long) cell.getNumericCellValue())
                }

            case CellType.BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue())

            case CellType.FORMULA:
                try {
                    return cell.getStringCellValue().trim()
                } catch (Exception e) {
                    try {
                        if (DateUtil.isCellDateFormatted(cell)) {
                            Date date = cell.getDateCellValue()
                            return new java.text.SimpleDateFormat("dd/MM/yyyy").format(date)
                        }
                        return String.valueOf((long) cell.getNumericCellValue())
                    } catch (Exception ee) {
                        return null
                    }
                }

            case CellType.BLANK:
                return null

            default:
                return null
        }
    }
}
