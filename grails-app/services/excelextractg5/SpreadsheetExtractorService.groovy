package excelextractg5

import builders.dsl.spreadsheet.builder.api.RowDefinition
import builders.dsl.spreadsheet.builder.api.SheetDefinition
import builders.dsl.spreadsheet.builder.api.WorkbookDefinition
import builders.dsl.spreadsheet.builder.poi.PoiSpreadsheetBuilder
import grails.gorm.transactions.ReadOnly
import javax.servlet.http.HttpServletResponse

import static excelextractg5.MyStyles.*

class SpreadsheetExtractorService {

    def generate(HttpServletResponse response) {
        log.info("Generating spreadsheet")

        PoiSpreadsheetBuilder.create(response.outputStream).build {workbookDef ->
            apply MyStyles

            sheet('Mock') {
                // header
                row{
                    style HEADERS
                    cell 'Parameter'
                    cell 'Cheese'
                    cell 'Tree'
                    cell 'Space'
                    cell 'Comments'
                    cell 'Layby'
                }
                row{
                    cell {
                        style SUBHEADING
                        value "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do"
                        colspan 6
                    }
                }

                ['Q1', 'Q2', 'Q3', 'Q4'].each {aQuestion ->
                    row {
                        cell aQuestion
                    }
                }

                row{
                    cell {
                        style SUBHEADING
                        value "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do"
                        colspan 6
                    }
                }

                ['Qu1', 'Qu2', 'Qu3', 'Qu4'].each {aQuestion ->
                    row {
                        cell aQuestion
                    }
                }

                ['CC1', 'CC2', 'CC3', 'CC4'].eachWithIndex{ String entry, int i ->
                    row( 3 + i ) {
                        cell(2) {
                            value entry
                        }
                    }
                }

                ['DI1', 'DI2', 'DI3', 'DI4'].eachWithIndex{ String entry, int i ->
                    row( 3 + i ) {
                        cell(3) {
                            value entry
                        }
                    }
                }

            }

            sheet('Abcd') {
                // header
                row {
                    style HEADERS                             // use the headers style defined above
                    cell 'Question'
                    cell 'Value'
                }
                fetchAbcds(1).each {anAbcd ->
                    row {
                        cell anAbcd.question
                        cell anAbcd.value
                    }
                }
            }
            sheet('Wxyz') {
                // header
                row {
                    style HEADERS
                    cell 'Question'
                    cell 'Value'
                }
                fetchWxyzs(1).each {anWxyz ->
                    row {
                        cell anWxyz.question
                        cell anWxyz.value
                    }
                }
            }

            sheet('Random') {
                filter auto             // automatic filters
                row{
                    cell 'Question'
                    cell 'Value'
                }
                row{
                    cell 'Why'
                    cell 'Because'
                }
                row{
                    cell 'Who'
                    cell 'Them'
                }

                row 5, { cell 'Line 5' }    // specify the row number
                row()                                                           // leave row 6 empty
                row { cell 'Line 7' }                       // start again on line 7

                row{
                    cell(2) {                               // specify the cell number, starting at 1
                        value 'Line 8 column (2 or B)'                   // use cell closure, call value
                        width 100                                           // width
                    }
                    cell('C') {                             // specify the cell letter, starting at A
                        formula('YEAR(TODAY())')                            // use cell formula
                    }
                }

                row {                                               // colspan and rowspan
                    cell {
                        value "Columns"
                        colspan 2
                    }
                }
                row {
                    cell {
                        value 'Rows'
                        rowspan 3
                    }
                    cell 'Value 1'
                }
                row {
                    cell ('B') { value 'Value 2' }
                }
                row {
                    cell ('B') { value 'Value 3' }
                }

                row {
                    cell('D') {
                        png image from 'grails-app/assets/images/apple-touch-icon.png'          // add an image
                    }
                }

                row {
                    cell 'A'
                    cell 'B'
                    cell 'A + B'
                }

                // naming cells allows you to refer to them later
                row {
                    cell {
                        value 10
                        name '_CellA'
                    }
                    cell {
                        value 20
                        name '_CellB'
                    }
                    cell {
                        formula 'SUM(#{_CellA},#{_CellB})'
                    }
                }
            }
            sheet('Links') {
                row {
                    cell {
                        value 'Go to Random sheet'
                        link to name '_CellA'
                        width auto
                    }
                    cell {
                        value 'File'
                        link to file 'text.txt'
                    }
                    cell {
                        value 'URL'
                        link to url 'https://www.bbc.com'
                    }
                    cell {
                        value 'Mail (plain)'
                        link to email 'someone@example.com'
                    }
                    cell {
                        value 'Mail (with subject)'
                        link to email 'someone@example.com',
                                cc: 'someoneelse@example.com'
                        subject: 'Testing Excel Builder'
                        body: 'It is really great tools'
                    }
                }
            }

            // pass the workbook to another method
            anotherSheet( workbookDef )

        }
    }

    /**
     * Pass a sheet to another method
     */
    def anotherSheet( WorkbookDefinition workbookDef ) {
        workbookDef.sheet( 'Another' ) {sheetDef ->
            addToSheet( sheetDef )
        }
    }

    /**
     * Pass a row to another method
     */
    def addToSheet( SheetDefinition sheetDef ) {
        sheetDef.row{
            cell( 'H1' )
            cell( 'H2' )
            cell( 'H3' )
            cell( 'H4' )
        }

        10.times { counter ->
            sheetDef.row { rowDef ->
                addRow( rowDef, counter )
            }
        }
    }

    /**
     * Work with row
     */
    def addRow( RowDefinition row, Integer rowNumber ) {
        row.cell(rowNumber)
        3.times {colIdx ->
            row.cell(UUID.randomUUID().toString().substring( 0, colIdx + rowNumber ))
        }
    }

    @ReadOnly
    List<Abcd> fetchAbcds( reqId ) {
        Abcd.findAllByReqId( reqId )
    }

    @ReadOnly
    List<Wxyz> fetchWxyzs( reqId ) {
        Wxyz.findAllByReqId( reqId )
    }

}
