package excelextractg5


import org.apache.poi.xssf.usermodel.XSSFWorkbookType
import org.springframework.http.HttpStatus

import java.time.LocalDateTime

class SpreadsheetExtractorController {

    def spreadsheetExtractorService

    def index() { }

    def generate() {
        response.setContentType(XSSFWorkbookType.XLSX.getContentType())
        response.setHeader "Content-disposition", "attachment; filename=myReport_${LocalDateTime.now()}.xlsx"

        try {
            spreadsheetExtractorService.generate(response)
        } catch(Exception e) {
            log.error("Error generating spreadsheet: ${e.message}", e)
            respond status: HttpStatus.INTERNAL_SERVER_ERROR.value()
        } finally {
            response.outputStream.flush()
            response.outputStream.close()
        }
    }
}
